// Feature module; shared bindings are imported explicitly.
import {
  state,
} from "../../app-core.js";
import {
  autoFilterDateSerial,
} from "./filter-and-text-tools.js";
import {
  queryTransformMedian,
} from "./query-transforms.js";
import {
  sheetUsedRange,
} from "./sort-and-autofilter.js";
import {
  decodeWorkbookRangeOrName,
  workbookCellReference,
} from "./validation-and-consolidation.js";
import {
  chartSvg,
  closeDataDialog,
  outlineRangeForAxis,
  uniqueSheetNameWithBase,
} from "../drawing/charts-and-insertion.js";
import {
  clampNumber,
  endingCellForPixelSize,
} from "../drawing/shape-rendering.js";
import {
  svgDataUrl,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  getDisplayForCell,
  materializeSelectionRangeForSheet,
  rangeToLabel,
  restoreGroupedSheetsAfterFailure,
  updateCellModel,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  pushHistory,
  pushSheetInsertHistory,
  restoreSheetFromHistory,
  storeCellModelForSheet,
} from "../editing/structure-and-history.js";
import {
  ensureCellModel,
  forEachCellInRange,
} from "../rendering/grid-and-canvas.js";
import {
  cellKey,
  columnName,
  columnWidthToPixels,
  dateToExcelSerial,
  escapeAttr,
  isDateFormat,
  rowHeightToPixels,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  applyExcelTableStyle,
  commitDataOperation,
  ensureSheetSize,
  quoteSheetNameForFormula,
  selectedDataRange,
} from "../ui/ribbon-and-links.js";
import {
  createBlankSheet,
  ensureSheetProtectionPermission,
  isSheetProtected,
  sheetProtectionAllows,
  showProtectedSheetMessage,
} from "../ui/sheet-protection-and-translation.js";
import {
  buildFormulaEngine,
  renderWorkbook,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  activeSheet,
  displayValue,
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  clampOutlineLevel,
  normalizeOutlineSettings,
} from "../workbook/tables-and-validation.js";


function dataAnalysisSeries(sheet, range, options) {
  const sheetId = getSheetId(sheet.name);
  const series = [];
  const byRows = options.groupedBy === "rows";
  const majorStart = byRows ? range.top : range.left;
  const majorEnd = byRows ? range.bottom : range.right;
  const minorStart = byRows ? range.left : range.top;
  const minorEnd = byRows ? range.right : range.bottom;
  for (let major = majorStart; major <= majorEnd; major += 1) {
    const labelCell = byRows
      ? getDisplayForCell(sheet, sheetId, major, range.left)
      : getDisplayForCell(sheet, sheetId, range.top, major);
    const label = options.labels && String(labelCell || "").trim()
      ? String(labelCell).trim()
      : byRows ? `${major}` : columnName(major);
    const start = minorStart + (options.labels ? 1 : 0);
    const values = [];
    const sourceIndexes = [];
    for (let minor = start; minor <= minorEnd; minor += 1) {
      const row = byRows ? major : minor;
      const col = byRows ? minor : major;
      const value = dataAnalysisCellNumericValue(sheet, sheetId, row, col);
      if (Number.isFinite(value)) {
        values.push(value);
        sourceIndexes.push(byRows ? col : row);
      }
    }
    if (values.length) series.push({ label, values, sourceIndexes });
  }
  return series;
}



function dataAnalysisCellNumericValue(sheet, sheetId, row, col) {
  const raw = sheet?.cells?.[cellKey(row, col)]?.raw;
  const value = raw != null ? raw : getDisplayForCell(sheet, sheetId, row, col);
  if (value == null) return null;
  if (typeof value === "string" && value.trim() === "") return null;
  const number = typeof value === "number" ? value : Number(String(value).replace(/[,¥$€£%]/g, ""));
  return Number.isFinite(number) ? number : null;
}



function dataAnalysisInputMatrix(sheet, range, options) {
  const sheetId = getSheetId(sheet.name);
  const top = range.top + (options.labels ? 1 : 0);
  const left = range.left + (options.labels ? 1 : 0);
  const rowLabels = [];
  const colLabels = [];
  if (options.labels) {
    for (let row = top; row <= range.bottom; row += 1) {
      rowLabels.push(String(getDisplayForCell(sheet, sheetId, row, range.left) || `行 ${row - top + 1}`));
    }
    for (let col = left; col <= range.right; col += 1) {
      colLabels.push(String(getDisplayForCell(sheet, sheetId, range.top, col) || columnName(col)));
    }
  }
  const values = [];
  for (let row = top; row <= range.bottom; row += 1) {
    const rowValues = [];
    for (let col = left; col <= range.right; col += 1) {
      const value = dataAnalysisCellNumericValue(sheet, sheetId, row, col);
      rowValues.push(Number.isFinite(value) ? value : NaN);
    }
    if (rowValues.some(Number.isFinite)) values.push(rowValues);
  }
  return {
    values,
    rowLabels: rowLabels.length ? rowLabels : values.map((_row, index) => `行 ${index + 1}`),
    colLabels: colLabels.length ? colLabels : (values[0] || []).map((_value, index) => columnName(left + index)),
  };
}



function dataAnalysisDescriptiveRows(series, options) {
  const rows = [["", ...series.map((item) => item.label)]];
  const metrics = [];
  if (options.summaryStats) {
    metrics.push(
      ["平均", (values) => dataAnalysisMean(values)],
      ["標準誤差", (values) => dataAnalysisStdevS(values) / Math.sqrt(values.length)],
      ["中央値", dataAnalysisMedian],
      ["最頻値", dataAnalysisMode],
      ["標準偏差", dataAnalysisStdevS],
      ["標本分散", dataAnalysisVarianceS],
      ["尖度", dataAnalysisKurtosis],
      ["歪度", dataAnalysisSkewness],
      ["範囲", (values) => Math.max(...values) - Math.min(...values)],
      ["最小", (values) => Math.min(...values)],
      ["最大", (values) => Math.max(...values)],
      ["合計", (values) => values.reduce((sum, value) => sum + value, 0)],
      ["データの個数", (values) => values.length],
    );
  }
  if (options.confidenceLevelEnabled) {
    metrics.push([`信頼区間(${dataAnalysisConfidencePercentLabel(options.confidence, true)})`, (values) => dataAnalysisConfidenceInterval(values, options.confidence)]);
  }
  if (options.kthLargestEnabled && options.kthLargest) {
    metrics.push([`最大(${options.kthLargest})`, (values) => dataAnalysisKthValue(values, options.kthLargest, "largest")]);
  }
  if (options.kthSmallestEnabled && options.kthSmallest) {
    metrics.push([`最小(${options.kthSmallest})`, (values) => dataAnalysisKthValue(values, options.kthSmallest, "smallest")]);
  }
  if (!metrics.length) metrics.push(["データの個数", (values) => values.length]);
  metrics.forEach(([label, fn]) => rows.push([label, ...series.map((item) => dataAnalysisRound(fn(item.values)))]));
  return { valid: true, rows };
}



function dataAnalysisMatrixRows(title, series, calculator) {
  if (series.length < 2) return { valid: false, message: `${title}には数値系列が 2 個以上必要です。` };
  const rows = [["", ...series.map((item) => item.label)]];
  series.forEach((left, rowIndex) => {
    rows.push([
      left.label,
      ...series.map((right, colIndex) => {
        if (colIndex > rowIndex) return "";
        return dataAnalysisRound(calculator(left, right));
      }),
    ]);
  });
  return { valid: true, rows };
}



function dataAnalysisHistogramRows(sheet, options, series) {
  const values = series.flatMap((item) => item.values).filter(Number.isFinite).sort((a, b) => a - b);
  if (!values.length) return { valid: false, message: "ヒストグラムに使用する数値がありません。" };
  let bins = [];
  const binRange = decodeWorkbookRangeOrName(options.binText, sheet);
  if (binRange) {
    const binSheet = state.model?.sheets?.[binRange.sheetIndex ?? state.activeSheetIndex] || sheet;
    const binSheetId = getSheetId(binSheet.name);
    forEachCellInRange(binRange, (row, col) => {
      const value = dataAnalysisCellNumericValue(binSheet, binSheetId, row, col);
      if (Number.isFinite(value)) bins.push(value);
    });
    bins = [...new Set(bins)].sort((a, b) => a - b);
  }
  if (!bins.length) {
    const min = values[0];
    const max = values.at(-1);
    const step = (max - min || 1) / 10;
    bins = Array.from({ length: 10 }, (_item, index) => min + step * (index + 1));
  }
  const points = [];
  let previous = -Infinity;
  bins.forEach((bin) => {
    const count = values.filter((value) => value <= bin && value > previous).length;
    points.push({ label: dataAnalysisRound(bin), count });
    previous = bin;
  });
  const overflow = values.filter((value) => value > previous).length;
  points.push({ label: "その他", count: overflow });
  const orderedPoints = options.pareto
    ? [...points].sort((left, right) => right.count - left.count || String(left.label).localeCompare(String(right.label), "ja-JP"))
    : points;
  const rows = [["ビン", "頻度"]];
  if (options.cumulativePercent) rows[0].push("累積 %");
  let consumed = 0;
  orderedPoints.forEach((point) => {
    consumed += point.count;
    const row = [point.label, point.count];
    if (options.cumulativePercent) row.push(dataAnalysisRound(consumed / values.length));
    rows.push(row);
  });
  const charts = options.chartOutput ? [{
    type: "column",
    title: "ヒストグラム",
    chart: {
      labels: orderedPoints.map((point) => String(point.label)),
      series: [
        { name: "頻度", values: orderedPoints.map((point) => point.count), color: "#4472C4" },
        ...(options.cumulativePercent
          ? [{ name: "累積 %", values: orderedPoints.map((_point, index) => Number(rows[index + 1]?.[2]) * 100), color: "#ED7D31" }]
          : []),
      ],
    },
  }] : [];
  return { valid: true, rows, charts };
}



function dataAnalysisMovingAverageRows(series, options) {
  const interval = Math.max(2, options.interval);
  const headers = ["期"];
  series.forEach((item) => {
    headers.push(`${item.label} 実績`);
    headers.push(`${item.label} 移動平均`);
    if (options.standardErrorOutput) headers.push(`${item.label} 標準誤差`);
  });
  const rows = [headers];
  const maxLength = Math.max(...series.map((item) => item.values.length));
  const forecastSeries = series.map((item) => {
    const forecasts = Array(item.values.length).fill("");
    for (let index = interval - 1; index < item.values.length; index += 1) {
      forecasts[index] = dataAnalysisMean(item.values.slice(index - interval + 1, index + 1));
    }
    return forecasts;
  });
  for (let index = 0; index < maxLength; index += 1) {
    const row = [index + 1];
    series.forEach((item, seriesIndex) => {
      const actual = item.values[index];
      const average = forecastSeries[seriesIndex][index];
      row.push(Number.isFinite(actual) ? dataAnalysisRound(actual) : "");
      row.push(Number.isFinite(average) ? dataAnalysisRound(average) : "");
      if (options.standardErrorOutput) {
        row.push(dataAnalysisCumulativeStandardError(item.values, forecastSeries[seriesIndex], index));
      }
    });
    rows.push(row);
  }
  const charts = options.chartOutput ? [dataAnalysisChartFromOutputRows("移動平均", rows, "line")] : [];
  return { valid: true, rows, charts };
}



function dataAnalysisExponentialSmoothingRows(series, options) {
  const damping = clampNumber(options.alpha, 0.001, 0.999, 0.05);
  const headers = ["期"];
  series.forEach((item) => {
    headers.push(`${item.label} 実績`);
    headers.push(`${item.label} 予測値`);
    if (options.standardErrorOutput) headers.push(`${item.label} 標準誤差`);
  });
  const rows = [headers];
  const maxLength = Math.max(...series.map((item) => item.values.length));
  const forecasts = series.map((item) => {
    const out = [];
    item.values.forEach((value, index) => {
      if (index === 0) out[index] = "";
      else if (index === 1) out[index] = item.values[0];
      else out[index] = damping * out[index - 1] + (1 - damping) * item.values[index - 1];
    });
    return out;
  });
  for (let index = 0; index < maxLength; index += 1) {
    const row = [index + 1];
    forecasts.forEach((values, seriesIndex) => {
      const actual = series[seriesIndex].values[index];
      row.push(Number.isFinite(actual) ? dataAnalysisRound(actual) : "");
      row.push(Number.isFinite(values[index]) ? dataAnalysisRound(values[index]) : "");
      if (options.standardErrorOutput) {
        row.push(dataAnalysisCumulativeStandardError(series[seriesIndex].values, values, index));
      }
    });
    rows.push(row);
  }
  const charts = options.chartOutput ? [dataAnalysisChartFromOutputRows("指数平滑", rows, "line")] : [];
  return { valid: true, rows, charts };
}



function dataAnalysisCumulativeStandardError(actualValues, forecastValues, index) {
  const errors = [];
  for (let cursor = 0; cursor <= index; cursor += 1) {
    const actual = actualValues[cursor];
    const forecast = forecastValues[cursor];
    if (Number.isFinite(actual) && Number.isFinite(forecast)) errors.push(actual - forecast);
  }
  if (!errors.length) return "";
  return dataAnalysisRound(Math.sqrt(dataAnalysisMean(errors.map((value) => value ** 2))));
}



function dataAnalysisRegressionRows(series, options) {
  if (series.length < 2) return { valid: false, message: "回帰分析には Y 系列と X 系列が必要です。" };
  const xSeries = series.slice(1);
  const observations = dataAnalysisRegressionObservations(series, options);
  if (observations.length < xSeries.length + 2) return { valid: false, message: "回帰分析に必要なデータ数が不足しています。" };
  const xRows = [];
  const yRows = [];
  observations.forEach((observation) => {
    yRows.push(observation.y);
    xRows.push(options.constantZero ? observation.x : [1, ...observation.x]);
  });
  if (yRows.length < xSeries.length + 2) return { valid: false, message: "回帰分析に必要な有効な組み合わせデータが不足しています。" };
  const regression = dataAnalysisOls(yRows, xRows, { constantZero: options.constantZero });
  if (!regression) return { valid: false, message: "回帰分析を計算できませんでした。" };
  const coefficientLabels = options.constantZero ? xSeries.map((item) => item.label) : ["切片", ...xSeries.map((item) => item.label)];
  const critical = dataAnalysisStudentTInverseTwoTail(1 - options.confidence / 100, regression.dfResidual);
  const confidenceLabel = dataAnalysisConfidencePercentLabel(options.confidence);
  const confidenceLabelFixed = dataAnalysisConfidencePercentLabel(options.confidence, true);
  const rows = [
    ["概要"],
    ["回帰統計", ""],
    ["重相関 R", dataAnalysisRound(Math.sqrt(Math.max(0, regression.r2)))],
    ["重決定 R2", dataAnalysisRound(regression.r2)],
    ["補正 R2", dataAnalysisRound(regression.adjustedR2)],
    ["標準誤差", dataAnalysisRound(regression.standardError)],
    ["観測数", regression.n],
    [""],
    ["分散分析表"],
    ["", "自由度", "変動", "分散", "F", "有意 F"],
    ["回帰", regression.dfRegression, dataAnalysisRound(regression.ssRegression), dataAnalysisRound(regression.msRegression), dataAnalysisRound(regression.f), dataAnalysisRound(regression.significanceF)],
    ["残差", regression.dfResidual, dataAnalysisRound(regression.ssResidual), dataAnalysisRound(regression.msResidual), "", ""],
    ["合計", regression.dfTotal, dataAnalysisRound(regression.ssTotal), "", "", ""],
    [""],
    ["", "係数", "標準誤差", "t", "P-値", `下限 ${confidenceLabel}`, `上限 ${confidenceLabel}`, `下限 ${confidenceLabelFixed}`, `上限 ${confidenceLabelFixed}`],
    ...coefficientLabels.map((label, index) => [
      label,
      dataAnalysisRound(regression.coefficients[index]),
      dataAnalysisRound(regression.coefficientStdErrors[index]),
      dataAnalysisRound(regression.tStats[index]),
      dataAnalysisRound(regression.pValues[index]),
      dataAnalysisRound(regression.coefficients[index] - critical * regression.coefficientStdErrors[index]),
      dataAnalysisRound(regression.coefficients[index] + critical * regression.coefficientStdErrors[index]),
      dataAnalysisRound(regression.coefficients[index] - critical * regression.coefficientStdErrors[index]),
      dataAnalysisRound(regression.coefficients[index] + critical * regression.coefficientStdErrors[index]),
    ]),
    [""],
  ];
  if (options.residuals || options.standardizedResiduals) {
    rows.push(["残差出力"]);
    rows.push(["観測値", "予測値", "残差", ...(options.standardizedResiduals ? ["標準化残差"] : [])]);
    yRows.forEach((actual, index) => {
      const residual = regression.residuals[index];
      const row = [
        index + 1,
        dataAnalysisRound(regression.fitted[index]),
        dataAnalysisRound(residual),
      ];
      if (options.standardizedResiduals) {
        const leverage = clampNumber(regression.leverages?.[index], 0, 0.999999, 0);
        const denominator = regression.standardError * Math.sqrt(Math.max(0.000001, 1 - leverage));
        row.push(dataAnalysisRound(denominator ? residual / denominator : ""));
      }
      rows.push(row);
    });
  }
  const charts = [];
  if (options.chartOutput || options.lineFitPlot) {
    xSeries.forEach((xItem, xIndex) => {
      const xValues = xRows.map((row) => row[options.constantZero ? xIndex : xIndex + 1]);
      charts.push({
        type: "scatter",
        title: `${xItem.label} 線形近似プロット`,
        chart: {
          labels: xValues.map(String),
          axes: {
            category: { title: xItem.label },
            value: { title: series[0]?.label || "Y" },
          },
          series: [
            { name: "観測値", values: xValues, yValues: yRows, color: "#4472C4" },
            { name: "予測値", values: xValues, yValues: regression.fitted, color: "#ED7D31" },
          ],
        },
      });
    });
  }
  if (options.residualPlot) {
    xSeries.forEach((xItem, xIndex) => {
      const xValues = xRows.map((row) => row[options.constantZero ? xIndex : xIndex + 1]);
      charts.push({
        type: "scatter",
        title: `${xItem.label} 残差プロット`,
        chart: {
          labels: xValues.map(String),
          axes: {
            category: { title: xItem.label },
            value: { title: "残差" },
          },
          series: [{ name: "残差", values: xValues, yValues: regression.residuals, color: "#A5A5A5" }],
        },
      });
    });
  }
  if (options.normalProbabilityPlot) {
    const sortedY = yRows.slice().sort((left, right) => left - right);
    const probabilities = sortedY.map((_value, index) => dataAnalysisRound(((index + 0.5) / sortedY.length) * 100));
    charts.push({
      type: "scatter",
      title: "回帰分析: 正規確率",
      chart: {
        labels: probabilities.map((value) => `${value}%`),
        series: [{ name: series[0]?.label || "Y", values: sortedY, color: "#70AD47" }],
      },
    });
    rows.push(["正規確率出力"]);
    rows.push(["百分位", series[0]?.label || "Y"]);
    sortedY.forEach((value, index) => rows.push([probabilities[index], dataAnalysisRound(value)]));
  }
  return { valid: true, rows, charts };
}



function dataAnalysisRegressionObservations(series, options = {}) {
  const ySeries = series?.[0];
  const xSeries = (series || []).slice(1);
  if (!ySeries || !xSeries.length) return [];
  const allHaveIndexes = [ySeries, ...xSeries].every((item) => Array.isArray(item?.sourceIndexes) && item.sourceIndexes.length === item.values?.length);
  if (allHaveIndexes) {
    const xMaps = xSeries.map((item) => dataAnalysisSeriesValueMap(item));
    const observations = [];
    (ySeries.values || []).forEach((y, index) => {
      if (!Number.isFinite(y)) return;
      const key = ySeries.sourceIndexes[index];
      const x = xMaps.map((map) => map.get(key));
      if (x.every(Number.isFinite)) observations.push({ key, y, x });
    });
    return observations;
  }
  const y = ySeries.values || [];
  const n = Math.min(y.length, ...xSeries.map((item) => item.values?.length || 0));
  const observations = [];
  for (let index = 0; index < n; index += 1) {
    const x = xSeries.map((item) => item.values[index]);
    if (Number.isFinite(y[index]) && x.every(Number.isFinite)) observations.push({ key: index, y: y[index], x });
  }
  return observations;
}



function dataAnalysisAnovaRows(series, options = {}) {
  if (series.length < 2) return { valid: false, message: "分散分析には 2 個以上の系列が必要です。" };
  const all = series.flatMap((item) => item.values);
  const grandMean = dataAnalysisMean(all);
  const summaryRows = [["グループ", "標本数", "合計", "平均", "分散"]];
  series.forEach((item) => {
    summaryRows.push(dataAnalysisAnovaSummaryRow(item.label, item.values));
  });
  const ssBetween = series.reduce((sum, item) => sum + item.values.length * (dataAnalysisMean(item.values) - grandMean) ** 2, 0);
  const ssWithin = series.reduce((sum, item) => sum + item.values.reduce((inner, value) => inner + (value - dataAnalysisMean(item.values)) ** 2, 0), 0);
  const dfBetween = series.length - 1;
  const dfWithin = all.length - series.length;
  const msBetween = ssBetween / Math.max(1, dfBetween);
  const msWithin = ssWithin / Math.max(1, dfWithin);
  const fValue = msWithin ? msBetween / msWithin : "";
  return {
    valid: true,
    rows: [
      ["概要"],
      ...summaryRows,
      [""],
      ["分散分析表"],
      ["変動要因", "変動", "自由度", "分散", "F", "P-値", "F 境界値"],
      ["グループ間", dataAnalysisRound(ssBetween), dfBetween, dataAnalysisRound(msBetween), dataAnalysisRound(fValue), dataAnalysisRound(dataAnalysisFRightTail(fValue, dfBetween, dfWithin)), dataAnalysisRound(dataAnalysisFInverseRightTail(options.alpha, dfBetween, dfWithin))],
      ["グループ内", dataAnalysisRound(ssWithin), dfWithin, dataAnalysisRound(msWithin), "", "", ""],
      ["合計", dataAnalysisRound(ssBetween + ssWithin), dfBetween + dfWithin, "", "", "", ""],
    ],
  };
}



function dataAnalysisAnovaSummaryRow(label, values) {
  const clean = (values || []).filter(Number.isFinite);
  return [
    label,
    clean.length,
    dataAnalysisRound(dataAnalysisSum(clean)),
    dataAnalysisRound(dataAnalysisMean(clean)),
    clean.length > 1 ? dataAnalysisRound(dataAnalysisVarianceS(clean)) : "",
  ];
}



function dataAnalysisAnovaTwoFactorNoReplicationRows(sheet, range, options) {
  const matrix = dataAnalysisInputMatrix(sheet, range, options);
  const values = matrix.values.filter((row) => row.every(Number.isFinite));
  const rowCount = values.length;
  const colCount = values[0]?.length || 0;
  if (rowCount < 2 || colCount < 2) return { valid: false, message: "繰り返しのない二元配置には 2 行 x 2 列以上の完全な数値表が必要です。" };
  const total = values.flat().reduce((sum, value) => sum + value, 0);
  const n = rowCount * colCount;
  const correction = (total ** 2) / n;
  const ssTotal = values.flat().reduce((sum, value) => sum + value ** 2, 0) - correction;
  const rowTotals = values.map((row) => dataAnalysisSum(row));
  const colTotals = Array.from({ length: colCount }, (_item, col) => values.reduce((sum, row) => sum + row[col], 0));
  const ssRows = rowTotals.reduce((sum, value) => sum + (value ** 2) / colCount, 0) - correction;
  const ssCols = colTotals.reduce((sum, value) => sum + (value ** 2) / rowCount, 0) - correction;
  const ssError = Math.max(0, ssTotal - ssRows - ssCols);
  const dfRows = rowCount - 1;
  const dfCols = colCount - 1;
  const dfError = dfRows * dfCols;
  const msRows = ssRows / dfRows;
  const msCols = ssCols / dfCols;
  const msError = ssError / Math.max(1, dfError);
  const fRows = msError ? msRows / msError : NaN;
  const fCols = msError ? msCols / msError : NaN;
  const summary = [["概要", "データの個数", "合計", "平均", "分散"]];
  values.forEach((row, index) => {
    summary.push(dataAnalysisAnovaSummaryRow(matrix.rowLabels[index] || `行 ${index + 1}`, row));
  });
  summary.push([""]);
  for (let col = 0; col < colCount; col += 1) {
    const colValues = values.map((row) => row[col]).filter(Number.isFinite);
    summary.push(dataAnalysisAnovaSummaryRow(matrix.colLabels[col] || `列 ${col + 1}`, colValues));
  }
  return {
    valid: true,
    rows: [
      ...summary,
      [""],
      ["分散分析表"],
      ["変動要因", "変動", "自由度", "分散", "F", "P-値", "F 境界値"],
      ["行", dataAnalysisRound(ssRows), dfRows, dataAnalysisRound(msRows), dataAnalysisRound(fRows), dataAnalysisRound(dataAnalysisFRightTail(fRows, dfRows, dfError)), dataAnalysisRound(dataAnalysisFInverseRightTail(options.alpha, dfRows, dfError))],
      ["列", dataAnalysisRound(ssCols), dfCols, dataAnalysisRound(msCols), dataAnalysisRound(fCols), dataAnalysisRound(dataAnalysisFRightTail(fCols, dfCols, dfError)), dataAnalysisRound(dataAnalysisFInverseRightTail(options.alpha, dfCols, dfError))],
      ["誤差", dataAnalysisRound(ssError), dfError, dataAnalysisRound(msError), "", "", ""],
      ["合計", dataAnalysisRound(ssTotal), n - 1, "", "", "", ""],
    ],
  };
}



function dataAnalysisAnovaTwoFactorReplicationRows(sheet, range, options) {
  const matrix = dataAnalysisInputMatrix(sheet, range, options);
  const replications = Math.max(2, Math.round(options.interval || 2));
  const sourceRows = matrix.values.filter((row) => row.every(Number.isFinite));
  const colCount = sourceRows[0]?.length || 0;
  const sampleCount = Math.floor(sourceRows.length / replications);
  if (sampleCount < 2 || colCount < 2) return { valid: false, message: "繰り返しのある二元配置には、各標本あたりの行数と 2 列以上の完全な数値表が必要です。" };
  const values = Array.from({ length: sampleCount }, (_sample, sampleIndex) =>
    Array.from({ length: colCount }, (_col, colIndex) =>
      Array.from({ length: replications }, (_rep, repIndex) => sourceRows[sampleIndex * replications + repIndex][colIndex]),
    ),
  );
  const allValues = values.flat(2);
  const grandMean = dataAnalysisMean(allValues);
  const sampleMeans = values.map((sample) => dataAnalysisMean(sample.flat()));
  const columnMeans = Array.from({ length: colCount }, (_item, col) => dataAnalysisMean(values.flatMap((sample) => sample[col])));
  const cellMeans = values.map((sample) => sample.map((cell) => dataAnalysisMean(cell)));
  const ssTotal = allValues.reduce((sum, value) => sum + (value - grandMean) ** 2, 0);
  const ssSamples = colCount * replications * sampleMeans.reduce((sum, mean) => sum + (mean - grandMean) ** 2, 0);
  const ssColumns = sampleCount * replications * columnMeans.reduce((sum, mean) => sum + (mean - grandMean) ** 2, 0);
  let ssInteraction = 0;
  let ssWithin = 0;
  for (let sample = 0; sample < sampleCount; sample += 1) {
    for (let col = 0; col < colCount; col += 1) {
      ssInteraction += replications * (cellMeans[sample][col] - sampleMeans[sample] - columnMeans[col] + grandMean) ** 2;
      values[sample][col].forEach((value) => {
        ssWithin += (value - cellMeans[sample][col]) ** 2;
      });
    }
  }
  const dfSamples = sampleCount - 1;
  const dfColumns = colCount - 1;
  const dfInteraction = dfSamples * dfColumns;
  const dfWithin = sampleCount * colCount * (replications - 1);
  const msSamples = ssSamples / dfSamples;
  const msColumns = ssColumns / dfColumns;
  const msInteraction = ssInteraction / Math.max(1, dfInteraction);
  const msWithin = ssWithin / Math.max(1, dfWithin);
  const summary = [["概要", ...matrix.colLabels.slice(0, colCount), "合計"]];
  values.forEach((sample, index) => {
    const flat = sample.flat();
    summary.push([matrix.rowLabels[index * replications] || `標本 ${index + 1}`]);
    summary.push(["データの個数", ...sample.map((cell) => cell.length), flat.length]);
    summary.push(["合計", ...sample.map((cell) => dataAnalysisRound(dataAnalysisSum(cell))), dataAnalysisRound(dataAnalysisSum(flat))]);
    summary.push(["平均", ...sample.map((cell) => dataAnalysisRound(dataAnalysisMean(cell))), dataAnalysisRound(dataAnalysisMean(flat))]);
    summary.push(["分散", ...sample.map((cell) => cell.length > 1 ? dataAnalysisRound(dataAnalysisVarianceS(cell)) : ""), flat.length > 1 ? dataAnalysisRound(dataAnalysisVarianceS(flat)) : ""]);
    if (index < values.length - 1) summary.push([""]);
  });
  return {
    valid: true,
    rows: [
      ...summary,
      [""],
      ["分散分析表"],
      ["変動要因", "変動", "自由度", "分散", "F", "P-値", "F 境界値"],
      ["標本", dataAnalysisRound(ssSamples), dfSamples, dataAnalysisRound(msSamples), dataAnalysisRound(msSamples / msWithin), dataAnalysisRound(dataAnalysisFRightTail(msSamples / msWithin, dfSamples, dfWithin)), dataAnalysisRound(dataAnalysisFInverseRightTail(options.alpha, dfSamples, dfWithin))],
      ["列", dataAnalysisRound(ssColumns), dfColumns, dataAnalysisRound(msColumns), dataAnalysisRound(msColumns / msWithin), dataAnalysisRound(dataAnalysisFRightTail(msColumns / msWithin, dfColumns, dfWithin)), dataAnalysisRound(dataAnalysisFInverseRightTail(options.alpha, dfColumns, dfWithin))],
      ["交互作用", dataAnalysisRound(ssInteraction), dfInteraction, dataAnalysisRound(msInteraction), dataAnalysisRound(msInteraction / msWithin), dataAnalysisRound(dataAnalysisFRightTail(msInteraction / msWithin, dfInteraction, dfWithin)), dataAnalysisRound(dataAnalysisFInverseRightTail(options.alpha, dfInteraction, dfWithin))],
      ["繰り返し内", dataAnalysisRound(ssWithin), dfWithin, dataAnalysisRound(msWithin), "", "", ""],
      ["合計", dataAnalysisRound(ssTotal), allValues.length - 1, "", "", "", ""],
    ],
  };
}



function dataAnalysisRankPercentileRows(series) {
  const cleanSeries = (series || [])
    .map((item) => ({ label: item.label || "入力", values: (item.values || []).filter(Number.isFinite) }))
    .filter((item) => item.values.length);
  if (!cleanSeries.length) return { valid: false, message: "順位と百分位数に使用する数値がありません。" };
  const rows = [["順位と百分位数"]];
  cleanSeries.forEach((item, seriesIndex) => {
    if (seriesIndex > 0) rows.push([""]);
    rows.push(["点", item.label, "順位", "百分率"]);
    const sorted = item.values.slice().sort((a, b) => b - a);
    item.values.forEach((value, index) => {
      const rank = sorted.findIndex((candidate) => candidate <= value) + 1;
      const percentile = item.values.length <= 1 ? 1 : 1 - (rank - 1) / (item.values.length - 1);
      rows.push([index + 1, dataAnalysisRound(value), rank, dataAnalysisRound(percentile)]);
    });
  });
  return { valid: true, rows };
}



function dataAnalysisSamplingRows(series, options) {
  const cleanSeries = (series || []).map((item) => ({
    label: item.label,
    values: (item.values || []).filter(Number.isFinite),
  })).filter((item) => item.values.length);
  if (!cleanSeries.length) return { valid: false, message: "サンプリングする数値がありません。" };
  const interval = Math.max(1, options.interval);
  const sampleCount = Math.max(1, Math.round(Number(options.sampleCount) || 10));
  const rng = dataAnalysisRandomGenerator(options.randomSeed);
  const methodLabel = options.samplingMethod === "random" ? "ランダム サンプリング" : "周期サンプリング";
  const headers = ["標本番号", ...cleanSeries.map((item) => item.label || "入力")];
  const rows = [headers];
  if (options.samplingMethod === "random") {
    const sampledBySeries = cleanSeries.map((item) => {
      const pool = item.values.map((value, index) => ({ index, value }));
      const sampled = [];
      while (pool.length && sampled.length < sampleCount) {
        const pick = Math.floor(rng() * pool.length);
        sampled.push(pool.splice(pick, 1)[0]);
      }
      return sampled;
    });
    const height = Math.max(...sampledBySeries.map((items) => items.length));
    for (let index = 0; index < height; index += 1) {
      rows.push([
        index + 1,
        ...sampledBySeries.map((items) => {
          const item = items[index];
          return item ? dataAnalysisRound(item.value) : "";
        }),
      ]);
    }
  } else {
    const maxLength = Math.max(...cleanSeries.map((item) => item.values.length));
    const sampleIndexes = [];
    for (let index = interval - 1; index < maxLength && sampleIndexes.length < sampleCount; index += interval) {
      sampleIndexes.push(index);
    }
    sampleIndexes.forEach((sourceIndex) => {
      rows.push([
        sourceIndex + 1,
        ...cleanSeries.map((item) => {
          const value = item.values[sourceIndex];
          return Number.isFinite(value) ? dataAnalysisRound(value) : "";
        }),
      ]);
    });
  }
  return {
    valid: true,
    rows: [[methodLabel], ...rows],
  };
}



function dataAnalysisRandomRows(options) {
  const rng = dataAnalysisRandomGenerator(options.randomSeed);
  const discreteDistribution = options.randomDistribution === "discrete" ? dataAnalysisDiscreteDistribution(options) : null;
  if (options.randomDistribution === "discrete" && !discreteDistribution?.length) {
    return { valid: false, message: "離散分布には、値と確率を 2 列で指定した入力範囲が必要です。" };
  }
  const rows = [[`乱数 (${dataAnalysisRandomDistributionLabel(options.randomDistribution)})`]];
  const pattern = options.randomDistribution === "pattern" ? dataAnalysisPatternValues(options) : null;
  let patternIndex = 0;
  for (let row = 0; row < options.randomRows; row += 1) {
    rows.push(Array.from({ length: options.randomCols }, () => {
      if (pattern?.length) {
        const value = pattern[patternIndex % pattern.length];
        patternIndex += 1;
        return dataAnalysisRound(value);
      }
      const value = dataAnalysisRandomValue(options, rng, discreteDistribution);
      return typeof value === "number" ? dataAnalysisRound(value) : value;
    }));
  }
  return { valid: true, rows };
}



function dataAnalysisRandomValue(options, rng = Math.random, discreteDistribution = null) {
  const p = clampNumber(options.mean, 0, 1, 0.5);
  const trials = Math.max(1, Math.round(options.stddev || 1));
  switch (options.randomDistribution) {
    case "uniform":
      return options.lowerBound + rng() * (options.upperBound - options.lowerBound);
    case "bernoulli":
      return rng() < p ? 1 : 0;
    case "binomial": {
      let successes = 0;
      for (let index = 0; index < trials; index += 1) if (rng() < p) successes += 1;
      return successes;
    }
    case "poisson":
      return dataAnalysisRandomPoisson(Math.max(0.000001, options.mean || 1), rng);
    case "discrete":
      return dataAnalysisRandomDiscrete(discreteDistribution, rng);
    case "normal":
    default:
      return dataAnalysisRandomNormal(options.mean, options.stddev, rng);
  }
}



function dataAnalysisRandomDistributionLabel(value) {
  switch (value) {
    case "uniform": return "一様";
    case "bernoulli": return "ベルヌーイ";
    case "binomial": return "二項";
    case "poisson": return "ポアソン";
    case "pattern": return "パターン";
    case "discrete": return "離散";
    case "normal":
    default:
      return "正規";
  }
}



function dataAnalysisPatternValues(options) {
  const start = Number(options.lowerBound);
  const end = Number(options.upperBound);
  const step = Math.max(0.000001, Math.abs(Number(options.patternStep) || 1));
  if (!Number.isFinite(start) || !Number.isFinite(end)) return [0];
  const values = [];
  const direction = start <= end ? 1 : -1;
  const limit = 100000;
  let value = start;
  for (let count = 0; count < limit; count += 1) {
    values.push(value);
    value += step * direction;
    if ((direction > 0 && value > end) || (direction < 0 && value < end)) break;
  }
  return values.length ? values : [start];
}



function dataAnalysisDiscreteDistribution(options) {
  const range = decodeWorkbookRangeOrName(options.probabilityRangeText, activeSheet());
  if (!range) return [];
  const sheet = state.model?.sheets?.[range.sheetIndex ?? state.activeSheetIndex];
  const sheetId = getSheetId(sheet?.name);
  if (!sheet || sheetId == null || range.right - range.left < 1) return [];
  const items = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    const valueCell = sheet.cells[cellKey(row, range.left)];
    const probability = dataAnalysisCellNumericValue(sheet, sheetId, row, range.left + 1);
    if (!Number.isFinite(probability) || probability <= 0) continue;
    const value = valueCell?.raw ?? valueCell?.cached ?? getDisplayForCell(sheet, sheetId, row, range.left);
    items.push({ value: value === "" || value == null ? 0 : value, probability });
  }
  const total = items.reduce((sum, item) => sum + item.probability, 0);
  if (!total) return [];
  let cumulative = 0;
  return items.map((item) => {
    cumulative += item.probability / total;
    return { value: item.value, cumulative };
  });
}



function dataAnalysisRandomDiscrete(distribution, rng = Math.random) {
  if (!distribution?.length) return "";
  const pick = rng();
  const item = distribution.find((entry) => pick <= entry.cumulative) || distribution.at(-1);
  return item?.value ?? "";
}



function dataAnalysisRandomGenerator(seedValue) {
  const seedText = String(seedValue ?? "").trim();
  if (!seedText) return Math.random;
  let seed = Math.abs(Math.round(Number(seedText))) || 1;
  return () => {
    seed = (seed * 1664525 + 1013904223) >>> 0;
    return seed / 4294967296;
  };
}



function dataAnalysisRandomPoisson(lambda, rng = Math.random) {
  const limit = Math.exp(-lambda);
  let k = 0;
  let product = 1;
  do {
    k += 1;
    product *= rng();
  } while (product > limit && k < 10000);
  return k - 1;
}



function dataAnalysisChartFromOutputRows(title, rows, type = "column") {
  const header = rows[0] || [];
  const dataRows = rows.slice(1);
  const series = [];
  for (let col = 1; col < header.length; col += 1) {
    const values = dataRows.map((row) => Number(row[col])).map((value) => (Number.isFinite(value) ? value : null));
    if (!values.some((value) => Number.isFinite(value))) continue;
    series.push({ name: String(header[col] || `系列 ${col}`), values, color: col === 1 ? "#4472C4" : "#ED7D31" });
  }
  return {
    type,
    title,
    chart: {
      labels: dataRows.map((row) => String(row[0] ?? "")),
      series,
    },
  };
}



function dataAnalysisFTestRows(series, options = {}) {
  if (series.length < 2) return { valid: false, message: "F 検定には 2 個の系列が必要です。" };
  const a = series[0].values;
  const b = series[1].values;
  if (a.length < 2 || b.length < 2) return { valid: false, message: "F 検定には各系列に 2 個以上の数値が必要です。" };
  const varA = dataAnalysisVarianceS(a);
  const varB = dataAnalysisVarianceS(b);
  const f = varB ? varA / varB : "";
  const pOneTail = dataAnalysisFRightTail(f, a.length - 1, b.length - 1);
  return {
    valid: true,
    rows: [
      ["F 検定: 2 標本を使った分散の検定", series[0].label, series[1].label],
      ["平均", dataAnalysisRound(dataAnalysisMean(a)), dataAnalysisRound(dataAnalysisMean(b))],
      ["分散", dataAnalysisRound(varA), dataAnalysisRound(varB)],
      ["観測数", a.length, b.length],
      ["自由度", a.length - 1, b.length - 1],
      ["F", dataAnalysisRound(f), ""],
      ["P(F<=f) 片側", dataAnalysisRound(pOneTail), ""],
      ["F 境界値 片側", dataAnalysisRound(dataAnalysisFInverseRightTail(options.alpha, a.length - 1, b.length - 1)), ""],
    ],
  };
}



function dataAnalysisTTestRows(series, options, mode) {
  if (series.length < 2) return { valid: false, message: "t 検定には 2 個の系列が必要です。" };
  const pairedValues = mode === "paired" ? dataAnalysisPairedNumericValues(series[0], series[1]) : null;
  const a = mode === "paired" ? pairedValues.left : series[0].values;
  const b = mode === "paired" ? pairedValues.right : series[1].values;
  if (mode === "paired" && Math.min(a.length, b.length) < 2) return { valid: false, message: "一対の t 検定には対応するデータが 2 組以上必要です。" };
  if (mode !== "paired" && (a.length < 2 || b.length < 2)) return { valid: false, message: "t 検定には各系列に 2 個以上の数値が必要です。" };
  const hypothesized = Number(options.hypothesizedMeanDiff) || 0;
  const meanA = dataAnalysisMean(a);
  const meanB = dataAnalysisMean(b);
  const varA = dataAnalysisVarianceS(a);
  const varB = dataAnalysisVarianceS(b);
  let t = "";
  let df = "";
  if (mode === "paired") {
    const n = Math.min(a.length, b.length);
    const diff = Array.from({ length: n }, (_item, index) => a[index] - b[index] - hypothesized);
    t = dataAnalysisMean(diff) / (dataAnalysisStdevS(diff) / Math.sqrt(Math.max(1, n)));
    df = n - 1;
  } else if (mode === "equal") {
    const pooled = (((a.length - 1) * varA) + ((b.length - 1) * varB)) / Math.max(1, a.length + b.length - 2);
    t = (meanA - meanB - hypothesized) / Math.sqrt(pooled * (1 / a.length + 1 / b.length));
    df = a.length + b.length - 2;
  } else {
    const seA = varA / a.length;
    const seB = varB / b.length;
    t = (meanA - meanB - hypothesized) / Math.sqrt(seA + seB);
    df = ((seA + seB) ** 2) / ((seA ** 2) / Math.max(1, a.length - 1) + (seB ** 2) / Math.max(1, b.length - 1));
  }
  const pTwoTail = dataAnalysisStudentTTwoTail(t, df);
  const pOneTail = pTwoTail / 2;
  const tCriticalOneTail = dataAnalysisStudentTInverseTwoTail(options.alpha * 2, df);
  const tCriticalTwoTail = dataAnalysisStudentTInverseTwoTail(options.alpha, df);
  const rows = [
    [dataAnalysisTTestTitle(mode), series[0].label, series[1].label],
    ["平均", dataAnalysisRound(meanA), dataAnalysisRound(meanB)],
    ["分散", dataAnalysisRound(varA), dataAnalysisRound(varB)],
    ["観測数", a.length, b.length],
  ];
  if (mode === "paired") {
    const n = Math.min(a.length, b.length);
    rows.push(["ピアソン相関", dataAnalysisRound(dataAnalysisCorrelation(a.slice(0, n), b.slice(0, n))), ""]);
  }
  rows.push(
    ["仮説平均との差異", hypothesized, ""],
    ["自由度", dataAnalysisRound(df), ""],
    ["t", dataAnalysisRound(t), ""],
    ["P(T<=t) 片側", dataAnalysisRound(pOneTail), ""],
    ["t 境界値 片側", dataAnalysisRound(tCriticalOneTail), ""],
    ["P(T<=t) 両側", dataAnalysisRound(pTwoTail), ""],
    ["t 境界値 両側", dataAnalysisRound(tCriticalTwoTail), ""],
  );
  return {
    valid: true,
    rows,
  };
}



function dataAnalysisTTestTitle(mode) {
  if (mode === "paired") return "t 検定: 一対の標本による平均の検定";
  if (mode === "equal") return "t 検定: 等分散を仮定した 2 標本による検定";
  return "t 検定: 分散が等しくないと仮定した 2 標本による検定";
}



function dataAnalysisZTestRows(series, options) {
  if (series.length < 2) return { valid: false, message: "z 検定には 2 個の系列が必要です。" };
  const a = series[0].values;
  const b = series[1].values;
  if (!a.length || !b.length) return { valid: false, message: "z 検定には各系列に数値データが必要です。" };
  const meanA = dataAnalysisMean(a);
  const meanB = dataAnalysisMean(b);
  const hypothesized = Number(options.hypothesizedMeanDiff) || 0;
  const varianceA = Math.max(0.000001, Number(options.knownVariance1) || 1);
  const varianceB = Math.max(0.000001, Number(options.knownVariance2) || 1);
  const z = (meanA - meanB - hypothesized) / Math.sqrt(varianceA / a.length + varianceB / b.length);
  const pTwoTail = 2 * (1 - dataAnalysisNormalCdf(Math.abs(z)));
  const pOneTail = pTwoTail / 2;
  return {
    valid: true,
    rows: [
      ["z 検定: 2 標本による平均の検定", series[0].label, series[1].label],
      ["平均", dataAnalysisRound(meanA), dataAnalysisRound(meanB)],
      ["既知の分散", dataAnalysisRound(varianceA), dataAnalysisRound(varianceB)],
      ["観測数", a.length, b.length],
      ["仮説平均との差異", hypothesized, ""],
      ["z", dataAnalysisRound(z), ""],
      ["P(Z<=z) 片側", dataAnalysisRound(pOneTail), ""],
      ["z 境界値 片側", dataAnalysisRound(dataAnalysisInverseNormal(1 - options.alpha)), ""],
      ["P(Z<=z) 両側", dataAnalysisRound(pTwoTail), ""],
      ["z 境界値 両側", dataAnalysisRound(dataAnalysisInverseNormal(1 - options.alpha / 2)), ""],
    ],
  };
}



function dataAnalysisFourierInputValues(sheet, range, options = {}) {
  const sheetId = getSheetId(sheet?.name);
  if (sheetId == null) return [];
  const singleColumn = range.left === range.right;
  const singleRow = range.top === range.bottom;
  let top = range.top;
  let left = range.left;
  if (options.labels) {
    if (singleColumn || !singleRow) top += 1;
    else left += 1;
  }
  const values = [];
  for (let row = top; row <= range.bottom; row += 1) {
    for (let col = left; col <= range.right; col += 1) {
      const value = dataAnalysisComplexValueFromCell(sheet, sheetId, row, col);
      if (value) values.push(value);
    }
  }
  return values;
}



function dataAnalysisComplexValueFromCell(sheet, sheetId, row, col) {
  const cell = sheet?.cells?.[cellKey(row, col)];
  const candidates = [
    cell?.raw,
    cell?.cached,
    getDisplayForCell(sheet, sheetId, row, col),
  ];
  for (const candidate of candidates) {
    const parsed = dataAnalysisParseComplexValue(candidate);
    if (parsed) return parsed;
  }
  return null;
}



function dataAnalysisNormalizeComplexValue(value) {
  if (value && typeof value === "object" && Number.isFinite(value.real) && Number.isFinite(value.imag)) {
    return { real: Number(value.real), imag: Number(value.imag) };
  }
  return dataAnalysisParseComplexValue(value);
}



function dataAnalysisParseComplexValue(value) {
  if (typeof value === "number") return Number.isFinite(value) ? { real: value, imag: 0 } : null;
  const text = String(value ?? "")
    .trim()
    .replace(/[,\s￥¥$€£]/g, "")
    .replace(/j$/i, "i")
    .replace(/I$/g, "i");
  if (!text) return null;
  const realOnly = Number(text);
  if (Number.isFinite(realOnly)) return { real: realOnly, imag: 0 };
  if (!/i$/i.test(text)) return null;
  const body = text.slice(0, -1);
  if (body === "" || body === "+") return { real: 0, imag: 1 };
  if (body === "-") return { real: 0, imag: -1 };
  const pureImag = Number(body);
  if (Number.isFinite(pureImag)) return { real: 0, imag: pureImag };
  const numberPattern = "(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:e[+-]?\\d+)?";
  const complexMatch = body.match(new RegExp(`^([+-]?${numberPattern})([+-])(${numberPattern})?$`, "i"));
  if (!complexMatch) return null;
  const real = Number(complexMatch[1]);
  const imagMagnitude = complexMatch[3] ? Number(complexMatch[3]) : 1;
  if (!Number.isFinite(real) || !Number.isFinite(imagMagnitude)) return null;
  return { real, imag: complexMatch[2] === "-" ? -imagMagnitude : imagMagnitude };
}



function dataAnalysisFourierRows(inputValues, options = {}) {
  const sourceValues = Array.isArray(inputValues) && inputValues[0]?.values ? inputValues[0].values : inputValues;
  const values = (sourceValues || []).map(dataAnalysisNormalizeComplexValue).filter(Boolean);
  if (values.length < 2) return { valid: false, message: "フーリエ解析には 2 個以上の数値または複素数が必要です。" };
  if (!dataAnalysisIsPowerOfTwo(values.length)) {
    return { valid: false, message: "フーリエ解析の入力データ数は 2 のべき乗で指定してください。" };
  }
  const n = values.length;
  const rows = [["k", "出力", "実数部", "虚数部"]];
  for (let k = 0; k < n; k += 1) {
    let real = 0;
    let imag = 0;
    for (let t = 0; t < n; t += 1) {
      const angle = ((options.inverseFourier ? 2 : -2) * Math.PI * k * t) / n;
      const cos = Math.cos(angle);
      const sin = Math.sin(angle);
      real += values[t].real * cos - values[t].imag * sin;
      imag += values[t].real * sin + values[t].imag * cos;
    }
    if (options.inverseFourier) {
      real /= n;
      imag /= n;
    }
    const roundedReal = dataAnalysisCleanSignedZero(dataAnalysisRound(real));
    const roundedImag = dataAnalysisCleanSignedZero(dataAnalysisRound(imag));
    rows.push([k, dataAnalysisComplexText(roundedReal, roundedImag), roundedReal, roundedImag]);
  }
  return { valid: true, rows };
}



function dataAnalysisCleanSignedZero(value) {
  return Object.is(value, -0) || Math.abs(Number(value)) < 1e-12 ? 0 : value;
}



function dataAnalysisIsPowerOfTwo(value) {
  const n = Number(value);
  return Number.isInteger(n) && n > 0 && (n & (n - 1)) === 0;
}



function dataAnalysisComplexText(real, imag) {
  const r = Number(real);
  const i = Number(imag);
  if (!Number.isFinite(r) || !Number.isFinite(i)) return "";
  const cleanReal = Math.abs(r) < 1e-12 ? 0 : r;
  const cleanImag = Math.abs(i) < 1e-12 ? 0 : i;
  if (cleanImag === 0) return String(cleanReal);
  if (cleanReal === 0) return `${cleanImag}i`;
  return `${cleanReal}${cleanImag < 0 ? "" : "+"}${cleanImag}i`;
}



function dataAnalysisOls(y, xRows, options = {}) {
  const xt = dataAnalysisTranspose(xRows);
  const xtx = dataAnalysisMatrixMultiply(xt, xRows);
  const xtxInv = dataAnalysisMatrixInverse(xtx);
  if (!xtxInv) return null;
  const xty = dataAnalysisMatrixVectorMultiply(xt, y);
  const coefficients = dataAnalysisMatrixVectorMultiply(xtxInv, xty);
  const fitted = xRows.map((row) => row.reduce((sum, value, index) => sum + value * coefficients[index], 0));
  const leverages = xRows.map((row) => {
    const projected = dataAnalysisMatrixVectorMultiply(xtxInv, row);
    return row.reduce((sum, value, index) => sum + value * (projected[index] || 0), 0);
  });
  const meanY = dataAnalysisMean(y);
  const ssTotal = options.constantZero
    ? y.reduce((sum, value) => sum + value ** 2, 0)
    : y.reduce((sum, value) => sum + (value - meanY) ** 2, 0);
  const residuals = y.map((value, index) => value - fitted[index]);
  const ssResidual = residuals.reduce((sum, value) => sum + value ** 2, 0);
  const ssRegression = Math.max(0, ssTotal - ssResidual);
  const r2 = ssTotal ? 1 - ssResidual / ssTotal : 1;
  const n = y.length;
  const p = options.constantZero ? coefficients.length : coefficients.length - 1;
  const dfRegression = Math.max(1, p);
  const dfResidual = Math.max(1, n - coefficients.length);
  const dfTotal = Math.max(1, n - (options.constantZero ? 0 : 1));
  const msRegression = ssRegression / dfRegression;
  const msResidual = ssResidual / dfResidual;
  const f = msResidual ? msRegression / msResidual : NaN;
  const adjustedR2 = n > p + 1 ? 1 - (1 - r2) * (n - 1) / (n - p - 1) : r2;
  const standardError = Math.sqrt(msResidual);
  const coefficientStdErrors = coefficients.map((_coef, index) => Math.sqrt(Math.max(0, msResidual * (xtxInv[index]?.[index] || 0))));
  const tStats = coefficients.map((coef, index) => {
    const se = coefficientStdErrors[index];
    return se ? coef / se : NaN;
  });
  const pValues = tStats.map((t) => dataAnalysisStudentTTwoTail(t, dfResidual));
  return {
    coefficients,
    coefficientStdErrors,
    tStats,
    pValues,
    fitted,
    residuals,
    leverages,
    r2,
    adjustedR2,
    standardError,
    n,
    ssTotal,
    ssRegression,
    ssResidual,
    dfRegression,
    dfResidual,
    dfTotal,
    msRegression,
    msResidual,
    f,
    significanceF: dataAnalysisFRightTail(f, dfRegression, dfResidual),
  };
}



function dataAnalysisTranspose(matrix) {
  return matrix[0].map((_value, col) => matrix.map((row) => row[col]));
}



function dataAnalysisMatrixMultiply(a, b) {
  return a.map((row) => b[0].map((_value, col) => row.reduce((sum, value, index) => sum + value * b[index][col], 0)));
}



function dataAnalysisMatrixVectorMultiply(matrix, vector) {
  return matrix.map((row) => row.reduce((sum, value, index) => sum + value * vector[index], 0));
}



function dataAnalysisMatrixInverse(matrix) {
  const n = matrix.length;
  const augmented = matrix.map((row, rowIndex) => [
    ...row.map((value) => Number(value)),
    ...Array.from({ length: n }, (_item, colIndex) => (colIndex === rowIndex ? 1 : 0)),
  ]);
  for (let col = 0; col < n; col += 1) {
    let pivot = col;
    for (let row = col + 1; row < n; row += 1) {
      if (Math.abs(augmented[row][col]) > Math.abs(augmented[pivot][col])) pivot = row;
    }
    if (Math.abs(augmented[pivot][col]) < 1e-12) return null;
    [augmented[col], augmented[pivot]] = [augmented[pivot], augmented[col]];
    const divisor = augmented[col][col];
    for (let item = 0; item < n * 2; item += 1) augmented[col][item] /= divisor;
    for (let row = 0; row < n; row += 1) {
      if (row === col) continue;
      const factor = augmented[row][col];
      for (let item = 0; item < n * 2; item += 1) augmented[row][item] -= factor * augmented[col][item];
    }
  }
  return augmented.map((row) => row.slice(n));
}



function dataAnalysisSum(values) {
  return values.reduce((sum, value) => sum + value, 0);
}



function dataAnalysisMean(values) {
  return values.length ? dataAnalysisSum(values) / values.length : NaN;
}



function dataAnalysisKthValue(values, k, mode = "largest") {
  const sorted = values.slice().filter(Number.isFinite).sort((left, right) => mode === "smallest" ? left - right : right - left);
  const index = Math.max(0, Math.round(Number(k) || 1) - 1);
  return sorted[index] ?? "";
}



function dataAnalysisConfidencePercentLabel(confidence, fixedOneDecimal = false) {
  const value = clampNumber(confidence, 1, 99.9, 95);
  if (fixedOneDecimal) return `${value.toFixed(1)}%`;
  return `${Number.isInteger(value) ? String(value) : String(value)}%`;
}



function dataAnalysisVarianceS(values) {
  if (values.length < 2) return NaN;
  const mean = dataAnalysisMean(values);
  return values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (values.length - 1);
}



function dataAnalysisVarianceP(values) {
  if (!values.length) return NaN;
  const mean = dataAnalysisMean(values);
  return values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / values.length;
}



function dataAnalysisStdevS(values) {
  return Math.sqrt(dataAnalysisVarianceS(values));
}



function dataAnalysisMedian(values) {
  const sorted = values.slice().sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}



function dataAnalysisMode(values) {
  const counts = new Map();
  values.forEach((value) => counts.set(value, (counts.get(value) || 0) + 1));
  const best = [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0] - b[0])[0];
  return best && best[1] > 1 ? best[0] : "";
}



function dataAnalysisSkewness(values) {
  const n = values.length;
  if (n < 3) return NaN;
  const mean = dataAnalysisMean(values);
  const sd = dataAnalysisStdevS(values);
  return (n / ((n - 1) * (n - 2))) * values.reduce((sum, value) => sum + ((value - mean) / sd) ** 3, 0);
}



function dataAnalysisKurtosis(values) {
  const n = values.length;
  if (n < 4) return NaN;
  const mean = dataAnalysisMean(values);
  const sd = dataAnalysisStdevS(values);
  const fourth = values.reduce((sum, value) => sum + ((value - mean) / sd) ** 4, 0);
  return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * fourth - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
}



function dataAnalysisSeriesValueMap(series) {
  const map = new Map();
  (series?.values || []).forEach((value, index) => {
    const key = series?.sourceIndexes?.[index] ?? index;
    if (Number.isFinite(value) && !map.has(key)) map.set(key, value);
  });
  return map;
}



function dataAnalysisPairedNumericValues(left, right) {
  const leftValues = Array.isArray(left) ? left : left?.values || [];
  const rightValues = Array.isArray(right) ? right : right?.values || [];
  const leftIndexes = Array.isArray(left?.sourceIndexes) ? left.sourceIndexes : null;
  const rightIndexes = Array.isArray(right?.sourceIndexes) ? right.sourceIndexes : null;
  if (leftIndexes && rightIndexes && leftIndexes.length === leftValues.length && rightIndexes.length === rightValues.length) {
    const rightMap = dataAnalysisSeriesValueMap(right);
    const pairedLeft = [];
    const pairedRight = [];
    leftValues.forEach((value, index) => {
      const other = rightMap.get(leftIndexes[index]);
      if (Number.isFinite(value) && Number.isFinite(other)) {
        pairedLeft.push(value);
        pairedRight.push(other);
      }
    });
    return { left: pairedLeft, right: pairedRight };
  }
  const n = Math.min(leftValues.length, rightValues.length);
  const pairedLeft = [];
  const pairedRight = [];
  for (let index = 0; index < n; index += 1) {
    const leftValue = leftValues[index];
    const rightValue = rightValues[index];
    if (Number.isFinite(leftValue) && Number.isFinite(rightValue)) {
      pairedLeft.push(leftValue);
      pairedRight.push(rightValue);
    }
  }
  return { left: pairedLeft, right: pairedRight };
}



function dataAnalysisCovarianceP(left, right) {
  const paired = dataAnalysisPairedNumericValues(left, right);
  const leftValues = paired.left;
  const rightValues = paired.right;
  const n = Math.min(leftValues.length, rightValues.length);
  if (!n) return NaN;
  const meanLeft = dataAnalysisMean(leftValues);
  const meanRight = dataAnalysisMean(rightValues);
  let sum = 0;
  for (let index = 0; index < n; index += 1) sum += (leftValues[index] - meanLeft) * (rightValues[index] - meanRight);
  return sum / n;
}



function dataAnalysisCorrelation(left, right) {
  const paired = dataAnalysisPairedNumericValues(left, right);
  const leftValues = paired.left;
  const rightValues = paired.right;
  const n = Math.min(leftValues.length, rightValues.length);
  if (n < 2) return NaN;
  const covariance = dataAnalysisCovarianceP(leftValues, rightValues);
  const sdLeft = Math.sqrt(dataAnalysisVarianceP(leftValues));
  const sdRight = Math.sqrt(dataAnalysisVarianceP(rightValues));
  return sdLeft && sdRight ? covariance / (sdLeft * sdRight) : NaN;
}



function dataAnalysisConfidenceInterval(values, confidence) {
  const sd = dataAnalysisStdevS(values);
  if (!Number.isFinite(sd)) return NaN;
  const df = values.length - 1;
  if (df <= 0) return NaN;
  const alpha = 1 - clampNumber(confidence, 1, 99.9, 95) / 100;
  const t = dataAnalysisStudentTInverseTwoTail(alpha, df);
  return t * sd / Math.sqrt(values.length);
}



function dataAnalysisStudentTTwoTail(t, df) {
  const x = Number(t);
  const v = Number(df);
  if (!Number.isFinite(x) || !Number.isFinite(v) || v <= 0) return NaN;
  const z = v / (v + x * x);
  return Math.min(1, dataAnalysisRegularizedBeta(z, v / 2, 0.5));
}



function dataAnalysisStudentTInverseTwoTail(probability, df) {
  const target = clampNumber(probability, 0.000001, 0.999999, 0.05);
  const v = Number(df);
  if (!Number.isFinite(v) || v <= 0) return NaN;
  let low = 0;
  let high = 1;
  while (dataAnalysisStudentTTwoTail(high, v) > target && high < 1000000) high *= 2;
  for (let index = 0; index < 80; index += 1) {
    const mid = (low + high) / 2;
    if (dataAnalysisStudentTTwoTail(mid, v) > target) low = mid;
    else high = mid;
  }
  return high;
}



function dataAnalysisFRightTail(f, df1, df2) {
  const x = Number(f);
  const a = Number(df1);
  const b = Number(df2);
  if (!Number.isFinite(x) || !Number.isFinite(a) || !Number.isFinite(b) || x < 0 || a <= 0 || b <= 0) return NaN;
  const z = (a * x) / (a * x + b);
  return 1 - dataAnalysisRegularizedBeta(z, a / 2, b / 2);
}



function dataAnalysisFInverseRightTail(probability, df1, df2) {
  const target = clampNumber(probability, 0.000001, 0.999999, 0.05);
  const a = Number(df1);
  const b = Number(df2);
  if (!Number.isFinite(a) || !Number.isFinite(b) || a <= 0 || b <= 0) return NaN;
  let low = 0;
  let high = 1;
  while (dataAnalysisFRightTail(high, a, b) > target && high < 1000000) high *= 2;
  for (let index = 0; index < 80; index += 1) {
    const mid = (low + high) / 2;
    if (dataAnalysisFRightTail(mid, a, b) > target) low = mid;
    else high = mid;
  }
  return high;
}



function dataAnalysisRegularizedBeta(x, a, b) {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const bt = Math.exp(dataAnalysisLogGamma(a + b) - dataAnalysisLogGamma(a) - dataAnalysisLogGamma(b) + a * Math.log(x) + b * Math.log(1 - x));
  if (x < (a + 1) / (a + b + 2)) return (bt * dataAnalysisBetaFraction(x, a, b)) / a;
  return 1 - (bt * dataAnalysisBetaFraction(1 - x, b, a)) / b;
}



function dataAnalysisBetaFraction(x, a, b) {
  const maxIterations = 100;
  const epsilon = 3e-7;
  const fpmin = 1e-30;
  let qab = a + b;
  let qap = a + 1;
  let qam = a - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < fpmin) d = fpmin;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= maxIterations; m += 1) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < fpmin) d = fpmin;
    c = 1 + aa / c;
    if (Math.abs(c) < fpmin) c = fpmin;
    d = 1 / d;
    h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < fpmin) d = fpmin;
    c = 1 + aa / c;
    if (Math.abs(c) < fpmin) c = fpmin;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < epsilon) break;
  }
  return h;
}



function dataAnalysisLogGamma(z) {
  const coefficients = [676.5203681218851, -1259.1392167224028, 771.3234287776531, -176.6150291621406, 12.507343278686905, -0.13857109526572012, 9.984369578019572e-6, 1.5056327351493116e-7];
  if (z < 0.5) return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * z)) - dataAnalysisLogGamma(1 - z);
  let x = 0.9999999999998099;
  const adjusted = z - 1;
  coefficients.forEach((coefficient, index) => {
    x += coefficient / (adjusted + index + 1);
  });
  const t = adjusted + coefficients.length - 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (adjusted + 0.5) * Math.log(t) - t + Math.log(x);
}



function dataAnalysisNormalCdf(value) {
  return 0.5 * (1 + dataAnalysisErf(value / Math.SQRT2));
}



function dataAnalysisErf(value) {
  const sign = value < 0 ? -1 : 1;
  const x = Math.abs(value);
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  const t = 1 / (1 + p * x);
  const y = 1 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
  return sign * y;
}



function dataAnalysisInverseNormal(p) {
  const clamped = clampNumber(p, 0.000001, 0.999999, 0.975);
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const plow = 0.02425;
  const phigh = 1 - plow;
  if (clamped < plow) {
    const q = Math.sqrt(-2 * Math.log(clamped));
    return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  if (clamped > phigh) {
    const q = Math.sqrt(-2 * Math.log(1 - clamped));
    return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  const q = clamped - 0.5;
  const r = q * q;
  return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
}



function dataAnalysisRandomNormal(mean, stddev, rng = Math.random) {
  const u1 = Math.max(Number.MIN_VALUE, rng());
  const u2 = rng();
  return mean + stddev * Math.sqrt(-2 * Math.log(u1)) * Math.cos(2 * Math.PI * u2);
}



function dataAnalysisRound(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return value === "" ? "" : "";
  return Math.round(numeric * 1000000) / 1000000;
}



function scenarioSummaryValueMaps(scenarios, changingCells, resultCells) {
  const points = [...changingCells, ...resultCells];
  const current = new Map();
  points.forEach((point) => current.set(scenarioPointKey(point), scenarioPointDisplay(point)));
  const scenarioValues = new Map();
  if (!resultCells.length) return { current, scenarios: scenarioValues };
  const snapshots = (state.model?.sheets || []).map((sheet) => captureSheetForHistory(sheet));
  const activeBefore = state.activeSheetIndex;
  try {
    scenarios.forEach((scenario) => {
      restoreScenarioSummarySnapshots(snapshots);
      (scenario.cells || []).forEach((point) => {
        const sheetIndex = point.sheetIndex ?? scenario.sheetIndex ?? state.activeSheetIndex;
        const sheet = state.model.sheets[sheetIndex];
        const sheetId = getSheetId(sheet?.name);
        if (sheet && sheetId != null) writeCellContent(sheet, sheetId, point.row, point.col, scenarioValueForPoint(scenario, { ...point, sheetIndex }) ?? "");
      });
      state.hf = buildFormulaEngine(state.model);
      const map = new Map();
      resultCells.forEach((point) => map.set(scenarioPointKey(point), scenarioPointDisplay(point)));
      scenarioValues.set(scenario.id, map);
    });
  } finally {
    restoreScenarioSummarySnapshots(snapshots);
    state.activeSheetIndex = activeBefore;
    state.hf = buildFormulaEngine(state.model);
  }
  return { current, scenarios: scenarioValues };
}



function restoreScenarioSummarySnapshots(snapshots) {
  (snapshots || []).forEach((snapshot, index) => {
    const sheet = state.model?.sheets?.[index];
    if (sheet) restoreSheetFromHistory(sheet, snapshot);
  });
}



function scenarioPointDisplay(point) {
  const sheet = state.model?.sheets?.[point.sheetIndex];
  if (!sheet) return "";
  const sheetId = getSheetId(sheet.name);
  return getDisplayForCell(sheet, sheetId, point.row, point.col);
}



function scenarioPointKey(point) {
  return `${point.sheetIndex}:${cellKey(point.row, point.col)}`;
}



function scenarioPointValueKey(point) {
  const sheetIndex = point?.sheetIndex ?? state.activeSheetIndex;
  return `${sheetIndex}:${columnName(point.col)}${point.row}`;
}



function scenarioPointAddressKey(point) {
  return `${columnName(point.col)}${point.row}`;
}



function scenarioValueForPoint(scenario, point) {
  if (!scenario?.values || !point) return undefined;
  const valueKey = scenarioPointValueKey(point);
  const addressKey = scenarioPointAddressKey(point);
  return scenario.values[valueKey] ?? scenario.values[addressKey];
}



function setScenarioValueForPoint(values, point, value) {
  if (!values || !point) return;
  values[scenarioPointValueKey(point)] = value;
}



function scenarioPointLabel(point, baseSheetIndex = state.activeSheetIndex) {
  const sheetIndex = point?.sheetIndex ?? baseSheetIndex;
  const address = `${columnName(point.col)}${point.row}`;
  if (sheetIndex === baseSheetIndex) return address;
  return workbookCellReference(sheetIndex, point.row, point.col);
}



function scenarioSummaryAddress(sheet, point) {
  return `${quoteSheetNameForFormula(sheet?.name || "")}!${columnName(point.col)}${point.row}`;
}



function openForecastSheetDialog() {
  const context = selectedDataRange("予測シート", { expandCurrentRegion: true });
  if (!context) return;
  const defaults = forecastDialogDefaults(context.sheet, context.range);
  state.dataDialog = { type: "forecast-sheet", range: context.range };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog forecast-sheet-dialog" role="dialog" aria-modal="true" aria-labelledby="forecastSheetTitle" data-data-dialog="forecast-sheet">
      <div class="excel-mini-titlebar" id="forecastSheetTitle">予測ワークシートの作成</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="forecast-preview" aria-hidden="true">
          <span class="forecast-preview-axis"></span>
          <span class="forecast-preview-line"></span>
          <span class="forecast-preview-band"></span>
        </div>
        <fieldset class="data-dialog-fieldset">
          <legend>グラフの種類</legend>
          <label class="data-dialog-check"><input type="radio" name="forecastChartType" data-data-field="chartType" data-forecast-sync value="line" checked />折れ線グラフ</label>
          <label class="data-dialog-check"><input type="radio" name="forecastChartType" data-data-field="chartType" data-forecast-sync value="column" />縦棒グラフ</label>
        </fieldset>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">予測期間</span><input type="number" min="1" max="120" data-data-field="periods" data-forecast-sync value="5" /></label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">予測終了</span><input type="text" inputmode="decimal" data-data-field="forecastEnd" data-forecast-sync value="${escapeAttr(defaults.forecastEnd)}" /></label>
        <fieldset class="data-dialog-fieldset forecast-options-fieldset">
          <legend>オプション</legend>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">季節性</span><input type="number" min="0" max="365" step="1" data-data-field="seasonality" data-forecast-sync value="0" /></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">不足ポイント</span>
            <select data-data-field="missingData" data-forecast-sync>
              <option value="interpolate">補間</option>
              <option value="zero">0</option>
            </select>
          </label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">重複タイムライン値の集計</span>
            <select data-data-field="duplicateAggregate" data-forecast-sync>
              <option value="average">平均</option>
              <option value="sum">合計</option>
              <option value="count">データの個数</option>
              <option value="min">最小値</option>
              <option value="max">最大値</option>
              <option value="median">中央値</option>
            </select>
          </label>
          <label class="data-dialog-check forecast-confidence-row">
            <input type="checkbox" data-data-field="confidence" data-forecast-sync checked />
            <span>信頼区間を表示する</span>
            <input type="number" min="0" max="99" step="1" data-data-field="confidenceLevel" data-forecast-sync value="95" />
            <span>%</span>
          </label>
        </fieldset>
        <div class="data-dialog-summary forecast-summary" data-forecast-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">作成</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncForecastDialog(overlay.querySelector("[data-data-dialog='forecast-sheet']"));
}



function applyForecastSheetDialog(dialog) {
  const range = state.dataDialog?.range;
  if (!range) return;
  const preview = forecastDialogPreview(dialog);
  if (preview.warning) {
    showMessageBox(preview.message, { title: "Cht WebSheet" });
    syncForecastDialog(dialog);
    return;
  }
  const { periods, forecastEnd, chartType, confidence, confidenceLevel, seasonality, missingData, duplicateAggregate } = forecastDialogOptions(dialog);
  if (workbookHasForecastSheet()) {
    showMessageBox("このブックには既に予測ワークシートがあります。新しい予測ワークシートを作成しますか?", {
      title: "Cht WebSheet",
      tone: "warning",
      buttons: [
        {
          label: "作成",
          result: "create",
          callback: () => {
            closeDataDialog();
            createForecastSheetFromSelection({ range, periods, forecastEnd, chartType, confidence, confidenceLevel, seasonality, missingData, duplicateAggregate });
          },
        },
        { label: "キャンセル", result: "cancel" },
      ],
      cancelResult: "cancel",
    });
    return;
  }
  closeDataDialog();
  createForecastSheetFromSelection({ range, periods, forecastEnd, chartType, confidence, confidenceLevel, seasonality, missingData, duplicateAggregate });
}



function handleForecastDialogSync(event) {
  syncForecastDialog(event.currentTarget.closest("[data-data-dialog='forecast-sheet']"));
}



function syncForecastDialog(dialog) {
  if (!dialog) return;
  const confidence = Boolean(dialog.querySelector("[data-data-field='confidence']")?.checked);
  const confidenceLevel = dialog.querySelector("[data-data-field='confidenceLevel']");
  if (confidenceLevel) confidenceLevel.disabled = !confidence;
  const preview = forecastDialogPreview(dialog);
  const summary = dialog.querySelector("[data-forecast-summary]");
  if (summary) {
    summary.textContent = preview.message;
    summary.classList.toggle("is-warning", preview.warning);
  }
  const previewBox = dialog.querySelector(".forecast-preview");
  if (previewBox) previewBox.dataset.chartType = forecastDialogOptions(dialog).chartType;
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = Boolean(preview.warning);
}



function forecastDialogOptions(dialog) {
  const aggregate = dialog.querySelector("[data-data-field='duplicateAggregate']")?.value;
  const forecastEndText = dialog.querySelector("[data-data-field='forecastEnd']")?.value || "";
  return {
    periods: Math.max(1, Math.min(120, Number(dialog.querySelector("[data-data-field='periods']")?.value) || 5)),
    forecastEnd: forecastEndValueFromDialogText(forecastEndText, activeSheet(), state.dataDialog?.range),
    forecastEndText,
    chartType: dialog.querySelector("[data-data-field='chartType']:checked")?.value === "column" ? "column" : "line",
    confidence: Boolean(dialog.querySelector("[data-data-field='confidence']")?.checked),
    confidenceLevel: Math.max(0, Math.min(99, Number(dialog.querySelector("[data-data-field='confidenceLevel']")?.value) || 95)),
    seasonality: Math.max(0, Math.min(365, Math.round(Number(dialog.querySelector("[data-data-field='seasonality']")?.value) || 0))),
    missingData: dialog.querySelector("[data-data-field='missingData']")?.value === "zero" ? "zero" : "interpolate",
    duplicateAggregate: ["average", "sum", "count", "min", "max", "median"].includes(aggregate) ? aggregate : "average",
  };
}



function forecastDialogPreview(dialog) {
  const sheet = activeSheet();
  const range = state.dataDialog?.range;
  if (!sheet || !range) return { warning: true, message: "予測に使用する範囲がありません。" };
  const options = forecastDialogOptions(dialog);
  const points = forecastSeriesPointsForRange(sheet, range, options);
  if (points.length < 2) {
    return { warning: true, message: "予測を作成するには、2つ以上の数値データが必要です。" };
  }
  const periods = forecastPeriodsForOptions(points, options);
  const step = forecastStep(points);
  const last = points[points.length - 1].x;
  if (String(options.forecastEndText || "").trim() && !Number.isFinite(options.forecastEnd)) {
    return { warning: true, message: "予測終了には有効なタイムライン値を指定してください。" };
  }
  if (Number.isFinite(options.forecastEnd) && options.forecastEnd <= last) {
    return { warning: true, message: "予測終了には、最後のタイムライン値より後の値を指定してください。" };
  }
  const end = last + step * periods;
  const timelineMeta = forecastTimelineOutputMetadata(sheet, range);
  const duplicateCount = forecastDuplicateTimelineCount(sheet, range);
  const confidenceText = options.confidence ? `信頼区間 ${options.confidenceLevel}% を表示します。` : "信頼区間は表示しません。";
  const seasonalityText = options.seasonality ? ` 季節性: ${options.seasonality}。` : "";
  const duplicateText = duplicateCount ? ` 重複タイムライン値 ${duplicateCount} 件は ${forecastAggregateLabel(options.duplicateAggregate)} で集計します。` : "";
  return {
    warning: false,
    message: `${points.length} 個の履歴ポイントから ${periods} 期間を予測します。予測終了: ${forecastTimelineOutputLabel(end, timelineMeta)}。${confidenceText}${seasonalityText}${duplicateText}`,
  };
}



function forecastDuplicateTimelineCount(sheet, range) {
  if (!sheet || !range || range.right <= range.left) return 0;
  const sheetId = getSheetId(sheet.name);
  const counts = new Map();
  for (let row = range.top; row <= range.bottom; row += 1) {
    const xValue = sheet.cells[cellKey(row, range.left)]?.raw ?? getDisplayForCell(sheet, sheetId, row, range.left);
    const x = forecastTimelineValue(xValue);
    if (!Number.isFinite(x)) continue;
    const key = String(x);
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  return [...counts.values()].filter((count) => count > 1).reduce((total, count) => total + count, 0);
}



function forecastAggregateLabel(value) {
  return {
    average: "平均",
    sum: "合計",
    count: "データの個数",
    min: "最小値",
    max: "最大値",
    median: "中央値",
  }[value] || "平均";
}



function workbookHasForecastSheet() {
  return Boolean((state.model?.sheets || []).some((sheet) => sheet?.forecast));
}



function createForecastSheetFromSelection(options = {}) {
  const context = options.range ? { sheet: activeSheet(), range: options.range } : selectedDataRange("予測シート");
  if (!context) return;
  const { sheet, range } = context;
  const points = forecastSeriesPointsForRange(sheet, range, options);
  if (points.length < 2) {
    showMessageBox("予測を作成するには、2つ以上の数値データが必要です。", { title: "Cht WebSheet" });
    return;
  }
  const periods = forecastPeriodsForOptions(points, options);
  const forecast = forecastRowsForOptions(points, periods, options);
  const confidenceLevel = Math.max(0, Math.min(99, Number(options.confidenceLevel) || 95));
  const residualStd = Number(forecast.residualStd) || 0;
  const timelineMeta = forecastTimelineOutputMetadata(sheet, range);
  const finalName = uniqueSheetNameWithBase("Forecast", state.model.sheets.map((item) => item.name));
  const targetSheet = createBlankSheet(finalName);
  ensureSheetSize(targetSheet, forecast.length + 1, options.confidence === false ? 3 : 5);
  updateCellModel(targetSheet, 1, 1, "期間");
  updateCellModel(targetSheet, 1, 2, "実績");
  updateCellModel(targetSheet, 1, 3, "予測");
  if (options.confidence !== false) {
    updateCellModel(targetSheet, 1, 4, "信頼下限");
    updateCellModel(targetSheet, 1, 5, "信頼上限");
  }
  forecast.forEach((item, index) => {
    const timeline = forecastTimelineOutputValue(item.x, timelineMeta);
    updateCellModel(targetSheet, index + 2, 1, timeline.value);
    if (timeline.numFmt) {
      const timelineCell = targetSheet.cells[cellKey(index + 2, 1)];
      if (timelineCell) timelineCell.numFmt = timeline.numFmt;
    }
    updateCellModel(targetSheet, index + 2, 2, item.actual ?? "");
    updateCellModel(targetSheet, index + 2, 3, item.forecast ?? "");
    if (options.confidence !== false && item.forecast != null) {
      const spread = forecastConfidenceSpread(item.forecast, confidenceLevel, residualStd);
      updateCellModel(targetSheet, index + 2, 4, item.forecast - spread);
      updateCellModel(targetSheet, index + 2, 5, item.forecast + spread);
    }
  });
  applyExcelTableStyle(targetSheet, { sheetIndex: state.model.sheets.length, top: 1, left: 1, bottom: forecast.length + 1, right: options.confidence === false ? 3 : 5 });
  addForecastChartToSheet(targetSheet, forecast, { ...options, timelineMeta });
  targetSheet.forecast = {
    sourceSheet: sheet.name,
    sourceRange: { ...range },
    chartType: options.chartType || "line",
    periods,
    forecastEnd: Number.isFinite(Number(options.forecastEnd)) ? Number(options.forecastEnd) : null,
    confidence: options.confidence !== false,
    confidenceLevel,
    seasonality: Math.max(0, Math.min(365, Math.round(Number(options.seasonality) || 0))),
    missingData: options.missingData === "zero" ? "zero" : "interpolate",
    duplicateAggregate: options.duplicateAggregate || "average",
  };
  const activeSheetBefore = state.activeSheetIndex;
  const insertIndex = state.model.sheets.length;
  state.model.sheets.push(targetSheet);
  state.activeSheetIndex = insertIndex;
  state.hf = buildFormulaEngine(state.model);
  renderWorkbook();
  setStatus(`${finalName} に予測シートを作成しました。`);
  pushSheetInsertHistory(insertIndex, targetSheet, activeSheetBefore, insertIndex, `${finalName} 予測シート`);
  scheduleWorkbookPersistence();
}



function forecastTimelineOutputMetadata(sheet, range) {
  if (!sheet || !range || range.right <= range.left) return { type: "number", numFmt: "" };
  const sheetId = getSheetId(sheet.name);
  for (let row = range.top; row <= range.bottom; row += 1) {
    const cell = sheet.cells[cellKey(row, range.left)];
    const raw = cell?.raw ?? cell?.cached ?? getDisplayForCell(sheet, sheetId, row, range.left);
    const display = getDisplayForCell(sheet, sheetId, row, range.left);
    if (!Number.isFinite(forecastTimelineValue(raw))) continue;
    if (cell?.numFmt && isDateFormat(cell.numFmt)) {
      return { type: "date", numFmt: cell.numFmt || "yyyy/m/d" };
    }
    if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
      return { type: "date", numFmt: cell?.numFmt || "yyyy/m/d" };
    }
    const displayText = String(display ?? "").trim();
    if (displayText && !/^-?\d+(\.\d+)?$/.test(displayText.replace(/,/g, "")) && autoFilterDateSerial(displayText) != null) {
      return { type: "date", numFmt: cell?.numFmt || "yyyy/m/d" };
    }
  }
  return { type: "number", numFmt: "" };
}



function forecastTimelineOutputValue(value, meta = {}) {
  if (meta.type === "date") {
    return { value, numFmt: meta.numFmt || "yyyy/m/d" };
  }
  return { value, numFmt: "" };
}



function forecastTimelineOutputLabel(value, meta = {}) {
  if (meta.type === "date") return displayValue(value, meta.numFmt || "yyyy/m/d");
  return String(value);
}



function forecastSeriesPointsForRange(sheet, range, options = {}) {
  const sheetId = getSheetId(sheet.name);
  const rawPoints = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    const yValue = sheet.cells[cellKey(row, range.right)]?.raw ?? getDisplayForCell(sheet, sheetId, row, range.right);
    const y = forecastNumericValue(yValue);
    const xValue = range.right > range.left
      ? sheet.cells[cellKey(row, range.left)]?.raw ?? getDisplayForCell(sheet, sheetId, row, range.left)
      : rawPoints.length + 1;
    const x = range.right > range.left ? forecastTimelineValue(xValue) : Number(xValue);
    if (Number.isFinite(x) && Number.isFinite(y)) {
      rawPoints.push({ x, y, row });
    } else if (Number.isFinite(x) && options.missingData === "zero") {
      rawPoints.push({ x, y: 0, row });
    }
  }
  return normalizeForecastPoints(rawPoints, options);
}



function forecastNumericValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = String(value ?? "").trim();
  if (!text) return NaN;
  const numeric = Number(text.replace(/,/g, ""));
  return Number.isFinite(numeric) ? numeric : NaN;
}



function forecastTimelineValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return dateToExcelSerial(value);
  const text = String(value ?? "").trim();
  if (!text) return NaN;
  const numeric = Number(text.replace(/,/g, ""));
  if (Number.isFinite(numeric)) return numeric;
  const serialDay = autoFilterDateSerial(text);
  if (serialDay != null) return serialDay + 25569;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? NaN : dateToExcelSerial(parsed);
}



function normalizeForecastPoints(points, options = {}) {
  const buckets = new Map();
  (points || []).forEach((point) => {
    if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return;
    const key = String(point.x);
    if (!buckets.has(key)) buckets.set(key, { x: point.x, values: [] });
    buckets.get(key).values.push(point.y);
  });
  const aggregate = options.duplicateAggregate || "average";
  const normalized = [...buckets.values()]
    .map((bucket) => ({ x: bucket.x, y: forecastAggregateDuplicateValues(bucket.values, aggregate) }))
    .filter((point) => Number.isFinite(point.y))
    .sort((left, right) => left.x - right.x);
  return options.missingData === "zero" ? normalized : interpolateForecastMissingPoints(normalized);
}



function forecastAggregateDuplicateValues(values, aggregate = "average") {
  const numbers = (values || []).filter((value) => Number.isFinite(value));
  if (!numbers.length) return NaN;
  if (aggregate === "sum") return numbers.reduce((total, value) => total + value, 0);
  if (aggregate === "count") return numbers.length;
  if (aggregate === "min") return Math.min(...numbers);
  if (aggregate === "max") return Math.max(...numbers);
  if (aggregate === "median") return queryTransformMedian(numbers);
  return numbers.reduce((total, value) => total + value, 0) / numbers.length;
}



function interpolateForecastMissingPoints(points) {
  const sorted = (points || []).slice().sort((left, right) => left.x - right.x);
  if (sorted.length < 3) return sorted;
  const step = forecastStep(sorted);
  if (!Number.isFinite(step) || step === 0) return sorted;
  const result = [];
  for (let index = 0; index < sorted.length - 1; index += 1) {
    const current = sorted[index];
    const next = sorted[index + 1];
    result.push(current);
    const gap = Math.round((next.x - current.x) / step);
    if (gap <= 1 || gap > 1000) continue;
    for (let offset = 1; offset < gap; offset += 1) {
      const ratio = offset / gap;
      result.push({
        x: current.x + step * offset,
        y: current.y + (next.y - current.y) * ratio,
        interpolated: true,
      });
    }
  }
  result.push(sorted[sorted.length - 1]);
  return result;
}



function forecastStep(points) {
  if (!points?.length) return 1;
  const rawStep = points.length > 1 ? points[points.length - 1].x - points[points.length - 2].x : 1;
  return Number.isFinite(rawStep) && rawStep !== 0 ? rawStep : 1;
}



function forecastDialogDefaults(sheet, range) {
  const points = forecastSeriesPointsForRange(sheet, range);
  const step = forecastStep(points);
  const last = points.length ? points[points.length - 1].x : 1;
  const end = last + step * 5;
  const timelineMeta = forecastTimelineOutputMetadata(sheet, range);
  return {
    step,
    forecastEnd: timelineMeta.type === "date"
      ? forecastTimelineOutputLabel(end, timelineMeta)
      : Number.isInteger(end)
        ? String(end)
        : String(Math.round(end * 100) / 100),
  };
}



function forecastEndValueFromDialogText(value, sheet, range) {
  const text = String(value ?? "").trim();
  if (!text) return NaN;
  const numeric = Number(text.replace(/,/g, ""));
  if (Number.isFinite(numeric)) return numeric;
  const timelineMeta = forecastTimelineOutputMetadata(sheet, range);
  const serialDay = autoFilterDateSerial(text);
  if (serialDay != null) return serialDay + 25569;
  if (timelineMeta.type === "date") {
    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) return dateToExcelSerial(parsed);
  }
  return NaN;
}



function forecastPeriodsForOptions(points, options = {}) {
  const requestedEnd = Number(options.forecastEnd);
  if (Number.isFinite(requestedEnd) && points.length) {
    const step = forecastStep(points);
    const distance = (requestedEnd - points[points.length - 1].x) / step;
    if (distance > 0) return Math.max(1, Math.min(120, Math.ceil(distance)));
  }
  return Math.max(1, Math.min(120, Number(options.periods) || 5));
}



function forecastConfidenceSpread(value, confidenceLevel = 95, residualStd = 0) {
  const level = Math.max(0, Math.min(99, Number(confidenceLevel) || 95));
  const scale = level <= 0 ? 0 : level / 95;
  if (Number.isFinite(residualStd) && residualStd > 0) return residualStd * 1.96 * scale;
  return Math.max(0, Math.abs(Number(value) || 0) * 0.1 * scale);
}



function addForecastChartToSheet(sheet, forecast, options = {}) {
  const width = 520;
  const height = 320;
  const labels = forecast.map((item) => forecastTimelineOutputLabel(item.x, options.timelineMeta));
  const actualValues = forecast.map((item) => Number.isFinite(item.actual) ? item.actual : null);
  const forecastValues = forecast.map((item) => Number.isFinite(item.forecast) ? item.forecast : null);
  const lastActualIndex = actualValues.map((value, index) => Number.isFinite(value) ? index : -1).filter((index) => index >= 0).pop();
  if (Number.isFinite(lastActualIndex) && lastActualIndex >= 0 && !Number.isFinite(forecastValues[lastActualIndex])) {
    forecastValues[lastActualIndex] = actualValues[lastActualIndex];
  }
  const series = [
    { name: "実績", values: actualValues, color: "#4472C4", marker: { symbol: "circle", size: 5 } },
    { name: "予測", values: forecastValues, color: "#ED7D31", marker: { symbol: "circle", size: 5 } },
  ];
  if (options.confidence !== false) {
    const confidenceLevel = Math.max(0, Math.min(99, Number(options.confidenceLevel) || 95));
    series.push(
      { name: "信頼下限", values: forecast.map((item) => Number.isFinite(item.forecast) ? item.forecast - forecastConfidenceSpread(item.forecast, confidenceLevel, forecast.residualStd) : null), color: "#A9D18E", marker: { symbol: "none" } },
      { name: "信頼上限", values: forecast.map((item) => Number.isFinite(item.forecast) ? item.forecast + forecastConfidenceSpread(item.forecast, confidenceLevel, forecast.residualStd) : null), color: "#A9D18E", marker: { symbol: "none" } },
    );
  }
  const chart = {
    labels,
    series,
    legend: { position: "bottom" },
    axes: {
      category: { title: "期間" },
      value: { title: "値" },
    },
    chartArea: { fillColor: "#ffffff", outlineColor: "#b8c2cc", outlineWidth: 1 },
    plotArea: { fillColor: "#f8fafc", outlineColor: "none" },
  };
  const type = options.chartType === "column" ? "column" : "line";
  const src = svgDataUrl(chartSvg(chart, type, { width, height, title: "予測" }));
  const startRow = 2;
  const startCol = Math.max(7, (options.confidence === false ? 3 : 5) + 2);
  const end = endingCellForPixelSize(sheet, startRow, startCol, width, height);
  ensureSheetSize(sheet, Math.max(sheet.rowCount, end.row), Math.max(sheet.colCount, end.col));
  sheet.images = Array.isArray(sheet.images) ? sheet.images : [];
  sheet.images.push({
    id: uniqueWorksheetObjectId("chart"),
    name: "予測グラフ",
    src,
    mimeType: "image/svg+xml",
    extension: "svg",
    editAs: "oneCell",
    tl: {
      col: startCol - 1,
      row: startRow - 1,
      nativeCol: startCol - 1,
      nativeRow: startRow - 1,
      nativeColOff: 0,
      nativeRowOff: 0,
    },
    br: null,
    ext: { width, height },
    hyperlink: "",
    chart: {
      type,
      sourceSheet: sheet.name,
      sourceRange: `A1:${columnName(options.confidence === false ? 3 : 5)}${forecast.length + 1}`,
      forecast: true,
    },
  });
}



function forecastRowsForOptions(points, extraCount, options = {}) {
  const seasonality = Math.max(0, Math.min(365, Math.round(Number(options.seasonality) || 0)));
  if (seasonality > 1 && points.length >= seasonality * 2) {
    return seasonalLinearForecast(points, extraCount, seasonality);
  }
  return linearForecast(points, extraCount);
}



function linearRegressionForForecast(points) {
  const n = points.length;
  const sumX = points.reduce((total, point) => total + point.x, 0);
  const sumY = points.reduce((total, point) => total + point.y, 0);
  const sumXY = points.reduce((total, point) => total + point.x * point.y, 0);
  const sumXX = points.reduce((total, point) => total + point.x * point.x, 0);
  const slope = (n * sumXY - sumX * sumY) / Math.max(0.000001, n * sumXX - sumX * sumX);
  const intercept = sumY / n - slope * (sumX / n);
  return { slope, intercept };
}



function linearForecast(points, extraCount) {
  const { slope, intercept } = linearRegressionForForecast(points);
  const rows = points.map((point) => ({ x: point.x, actual: point.y, forecast: null }));
  const step = forecastStep(points);
  let x = points[points.length - 1].x;
  for (let index = 0; index < extraCount; index += 1) {
    x += step;
    rows.push({ x, actual: null, forecast: slope * x + intercept });
  }
  rows.residualStd = forecastResidualStandardDeviation(points, (point) => slope * point.x + intercept);
  rows.method = "linear";
  return rows;
}



function seasonalLinearForecast(points, extraCount, seasonality) {
  const { slope, intercept } = linearRegressionForForecast(points);
  const buckets = Array.from({ length: seasonality }, () => []);
  points.forEach((point, index) => {
    buckets[index % seasonality].push(point.y - (slope * point.x + intercept));
  });
  const seasonalOffsets = buckets.map((values) =>
    values.length ? values.reduce((total, value) => total + value, 0) / values.length : 0,
  );
  const predict = (point, index) => slope * point.x + intercept + seasonalOffsets[index % seasonality];
  const rows = points.map((point) => ({ x: point.x, actual: point.y, forecast: null }));
  const step = forecastStep(points);
  let x = points[points.length - 1].x;
  for (let offset = 0; offset < extraCount; offset += 1) {
    x += step;
    const index = points.length + offset;
    rows.push({ x, actual: null, forecast: slope * x + intercept + seasonalOffsets[index % seasonality] });
  }
  rows.residualStd = forecastResidualStandardDeviation(points, predict);
  rows.method = "seasonal";
  rows.seasonality = seasonality;
  return rows;
}



function forecastResidualStandardDeviation(points, predictor) {
  const errors = (points || [])
    .map((point, index) => point.y - Number(predictor(point, index)))
    .filter((value) => Number.isFinite(value));
  if (errors.length < 2) return 0;
  const sumSquares = errors.reduce((total, value) => total + value * value, 0);
  return Math.sqrt(sumSquares / Math.max(1, errors.length - 1));
}



function openOutlineGroupDialog(delta) {
  const context = selectedDataRange(delta > 0 ? "グループ化" : "グループ解除");
  if (!context) return;
  const { range } = context;
  let defaultAxis = outlineAxisForRange(range);
  if (!outlineAxisSheetAllows(context.sheet, "rows") && !outlineAxisSheetAllows(context.sheet, "columns")) {
    showProtectedSheetMessage("保護されたシートではアウトラインを変更できません。行または列の書式設定を許可してください。");
    return;
  }
  if (!outlineAxisSheetAllows(context.sheet, defaultAxis)) {
    defaultAxis = outlineAxisSheetAllows(context.sheet, "rows") ? "rows" : "columns";
  }
  state.dataDialog = { type: delta > 0 ? "outline-group" : "outline-ungroup", range, delta };
  closeDataDialog({ keepState: true });
  const title = delta > 0 ? "グループ化" : "グループ解除";
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog outline-dialog" role="dialog" aria-modal="true" aria-labelledby="outlineDialogTitle" data-data-dialog="${delta > 0 ? "outline-group" : "outline-ungroup"}">
      <div class="excel-mini-titlebar" id="outlineDialogTitle">${title}</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">${delta > 0 ? "グループ化" : "グループ解除"}する方向を選択してください。</div>
        <label class="data-dialog-check"><input type="radio" name="outlineAxis" data-data-field="axis" data-outline-sync value="rows" ${defaultAxis === "rows" ? "checked" : ""} ${outlineAxisSheetAllows(context.sheet, "rows") ? "" : "disabled"} />行</label>
        <label class="data-dialog-check"><input type="radio" name="outlineAxis" data-data-field="axis" data-outline-sync value="columns" ${defaultAxis === "columns" ? "checked" : ""} ${outlineAxisSheetAllows(context.sheet, "columns") ? "" : "disabled"} />列</label>
        <div class="data-dialog-summary outline-summary" data-outline-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncOutlineDialog(overlay.querySelector("[data-data-dialog]"));
}



function applyOutlineDialog(dialog) {
  const context = state.dataDialog;
  const axis = dialog.querySelector("[data-data-field='axis']:checked")?.value || "rows";
  const delta = Number(context?.delta || 1);
  const range = context?.range;
  if (!ensureOutlineAxisAllowed(activeSheet(), axis)) return;
  const validation = outlineDialogValidation(dialog);
  if (!validation.valid) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    syncOutlineDialog(dialog);
    return;
  }
  closeDataDialog();
  if (range) applyOutlineToSelection(delta, delta > 0 ? "グループ化しました。" : "グループ解除しました。", axis, range);
}



function handleOutlineDialogSync(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog]");
  if (!dialog) return;
  if (dialog.dataset.dataDialog === "outline-group" || dialog.dataset.dataDialog === "outline-ungroup") syncOutlineDialog(dialog);
  if (dialog.dataset.dataDialog === "outline-settings") syncOutlineSettingsDialog(dialog);
}



function syncOutlineDialog(dialog) {
  if (!dialog) return;
  const validation = outlineDialogValidation(dialog);
  const summary = dialog.querySelector("[data-outline-summary]");
  if (summary) {
    summary.textContent = validation.message;
    summary.classList.toggle("is-warning", !validation.valid);
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = !validation.valid;
}



function outlineDialogValidation(dialog) {
  const range = state.dataDialog?.range;
  const sheet = activeSheet();
  if (!sheet || !range) return { valid: false, message: "対象範囲を選択してください。" };
  const axis = dialog?.querySelector("[data-data-field='axis']:checked")?.value || "rows";
  const delta = Number(state.dataDialog?.delta || 1);
  if (!outlineAxisSheetAllows(sheet, axis)) {
    return { valid: false, message: `保護されたシートでは${axis === "columns" ? "列" : "行"}のアウトラインを変更できません。` };
  }
  const start = axis === "columns" ? range.left : range.top;
  const end = axis === "columns" ? range.right : range.bottom;
  const levels = [];
  for (let index = start; index <= end; index += 1) levels.push(outlineLevelAt(sheet, axis, index));
  const maxLevel = Math.max(0, ...levels);
  const minLevel = Math.min(...levels);
  const count = Math.max(0, end - start + 1);
  if (delta > 0 && maxLevel >= 8) return { valid: false, message: "アウトライン レベルは 8 までです。" };
  if (delta < 0 && maxLevel <= 0) return { valid: false, message: "選択範囲に解除できるアウトラインはありません。" };
  const nextMin = Math.max(0, Math.min(8, minLevel + delta));
  const nextMax = Math.max(0, Math.min(8, maxLevel + delta));
  const axisLabel = axis === "columns" ? "列" : "行";
  const actionLabel = delta > 0 ? "グループ化" : "グループ解除";
  return {
    valid: true,
    message: `${rangeToLabel(range)} の ${count.toLocaleString("ja-JP")} ${axisLabel}を${actionLabel}します。現在レベル ${minLevel}-${maxLevel}、変更後 ${nextMin}-${nextMax}。`,
  };
}



function openOutlineSettingsDialog() {
  if (!state.model) return;
  const sheet = activeSheet();
  const settings = normalizeOutlineSettings(sheet.outlineSettings);
  state.dataDialog = { type: "outline-settings" };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog outline-settings-dialog" role="dialog" aria-modal="true" aria-labelledby="outlineSettingsTitle" data-data-dialog="outline-settings">
      <div class="excel-mini-titlebar" id="outlineSettingsTitle">設定</div>
      <div class="excel-mini-body data-dialog-body">
        <fieldset class="data-dialog-fieldset">
          <legend>方向</legend>
          <label class="data-dialog-check"><input type="checkbox" data-data-field="summaryBelow" data-outline-settings-sync ${settings.summaryBelow ? "checked" : ""} />集計行を詳細データの下に表示する</label>
          <label class="data-dialog-check"><input type="checkbox" data-data-field="summaryRight" data-outline-settings-sync ${settings.summaryRight ? "checked" : ""} />集計列を詳細データの右に表示する</label>
        </fieldset>
        <fieldset class="data-dialog-fieldset">
          <legend>スタイル</legend>
          <label class="data-dialog-check"><input type="checkbox" data-data-field="autoStyle" data-outline-settings-sync ${settings.applyStyles ? "checked" : ""} />自動スタイル</label>
        </fieldset>
        <div class="data-dialog-summary outline-settings-summary" data-outline-settings-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncOutlineSettingsDialog(overlay.querySelector("[data-data-dialog='outline-settings']"));
}



function applyOutlineSettingsDialog(dialog) {
  const sheet = activeSheet();
  if (!sheet) {
    closeDataDialog();
    return;
  }
  const nextSettings = normalizeOutlineSettings({
    summaryBelow: Boolean(dialog.querySelector("[data-data-field='summaryBelow']")?.checked),
    summaryRight: Boolean(dialog.querySelector("[data-data-field='summaryRight']")?.checked),
    applyStyles: Boolean(dialog.querySelector("[data-data-field='autoStyle']")?.checked),
  });
  const currentSettings = normalizeOutlineSettings(sheet.outlineSettings);
  const autoStyle = nextSettings.applyStyles;
  closeDataDialog();
  if (applyOutlineSettingsForSheetGroup(nextSettings, autoStyle)) return;
  if (!autoStyle && currentSettings.summaryBelow === nextSettings.summaryBelow && currentSettings.summaryRight === nextSettings.summaryRight && currentSettings.applyStyles === nextSettings.applyStyles) {
    setStatus("アウトライン設定は変更されていません。");
    return;
  }
  const range = activeSelectionRange() || { sheetIndex: state.activeSheetIndex, top: 1, bottom: 1, left: 1, right: 1 };
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.outlineSettings = nextSettings;
  if (autoStyle) applyOutlineAutomaticStyles(sheet);
  commitDataOperation(sheet, sheetBefore, range, "アウトライン設定", "アウトライン設定を変更しました。", range);
}



function applyOutlineSettingsForSheetGroup(nextSettings, autoStyle) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const items = sheetIndexes
    .map((sheetIndex) => ({ sheetIndex, sheet: state.model?.sheets?.[sheetIndex] }))
    .filter((item) => item.sheet);
  if (!items.length) return true;
  const changed = items.some(({ sheet }) => {
    const current = normalizeOutlineSettings(sheet.outlineSettings);
    return autoStyle || current.summaryBelow !== nextSettings.summaryBelow || current.summaryRight !== nextSettings.summaryRight || current.applyStyles !== nextSettings.applyStyles;
  });
  if (!changed) {
    setStatus("アウトライン設定は変更されていません。");
    return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach(({ sheet }) => {
      sheet.outlineSettings = normalizeOutlineSettings(nextSettings);
      if (autoStyle) applyOutlineAutomaticStyles(sheet);
    });
    state.hf = buildFormulaEngine(state.model);
    const range = activeSelectionRange() || { sheetIndex: state.activeSheetIndex, top: 1, bottom: 1, left: 1, right: 1 };
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
      selectionAfterRanges: activeSelectionRanges().map((item) => ({ ...item })),
      label: `${items.length} シート アウトライン設定`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートでアウトライン設定を変更しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`アウトライン設定を変更できませんでした：${error.message || error}`);
  }
  return true;
}



function syncOutlineSettingsDialog(dialog) {
  if (!dialog) return;
  const sheet = activeSheet();
  const groups = sheet ? outlineGroups(sheet, "rows").length + outlineGroups(sheet, "columns").length : 0;
  const autoStyle = Boolean(dialog.querySelector("[data-data-field='autoStyle']")?.checked);
  const summary = dialog.querySelector("[data-outline-settings-summary]");
  if (summary) {
    summary.textContent = `${groups.toLocaleString("ja-JP")} 個のアウトライン グループがあります。${autoStyle ? "OK で集計行/列に自動スタイルを適用します。" : "方向設定だけを変更します。"}`;
    summary.classList.toggle("is-warning", false);
  }
}



function applyOutlineAutomaticStyles(sheet) {
  const used = sheetUsedRange(sheet);
  if (!used) return;
  const settings = normalizeOutlineSettings(sheet.outlineSettings);
  const summaryRows = new Set();
  const summaryCols = new Set();
  outlineGroups(sheet, "rows").forEach((group) => {
    const row = settings.summaryBelow ? group.end + 1 : group.start - 1;
    if (row >= used.top && row <= used.bottom) summaryRows.add(row);
  });
  outlineGroups(sheet, "columns").forEach((group) => {
    const col = settings.summaryRight ? group.end + 1 : group.start - 1;
    if (col >= used.left && col <= used.right) summaryCols.add(col);
  });
  summaryRows.forEach((row) => {
    for (let col = used.left; col <= used.right; col += 1) {
      const existing = sheet.cells[cellKey(row, col)];
      if (!existing) continue;
      const cell = ensureCellModel(sheet, row, col);
      cell.css.fontWeight = "700";
      storeCellModelForSheet(sheet, cellKey(row, col), cell);
    }
  });
  summaryCols.forEach((col) => {
    for (let row = used.top; row <= used.bottom; row += 1) {
      const existing = sheet.cells[cellKey(row, col)];
      if (!existing) continue;
      const cell = ensureCellModel(sheet, row, col);
      cell.css.fontWeight = "700";
      storeCellModelForSheet(sheet, cellKey(row, col), cell);
    }
  });
}



function autoOutlineSelection() {
  const context = selectedDataRange("自動アウトライン", { expandCurrentRegion: true, limit: Number.MAX_SAFE_INTEGER });
  if (!context) return;
  if (autoOutlineSelectionForSheetGroup(context)) return;
  const { sheet, range } = context;
  if (!ensureAutoOutlineAllowed(sheet)) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const settings = normalizeOutlineSettings(sheet.outlineSettings);
  sheet.outlineSettings = settings;
  let changed = false;
  changed = applyAutoOutlineRows(sheet, range, settings) || changed;
  changed = applyAutoOutlineColumns(sheet, range, settings) || changed;
  if (!changed) {
    showMessageBox("数式または集計行/列を検出できませんでした。", { title: "Cht WebSheet" });
    return;
  }
  commitDataOperation(sheet, sheetBefore, range, `${rangeToLabel(range)} 自動アウトライン`, "自動アウトラインを作成しました。", range);
}



function autoOutlineSelectionForSheetGroup(context) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !context?.range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const range = materializeSelectionRangeForSheet(sheet, context.range, sheetIndex);
    if (!range) continue;
    if (!ensureAutoOutlineAllowed(sheet)) return true;
    items.push({ sheetIndex, sheet, range });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    let changed = false;
    items.forEach(({ sheet, range }) => {
      const settings = normalizeOutlineSettings(sheet.outlineSettings);
      sheet.outlineSettings = settings;
      changed = applyAutoOutlineRows(sheet, range, settings) || changed;
      changed = applyAutoOutlineColumns(sheet, range, settings) || changed;
    });
    if (!changed) {
      restoreGroupedSheetsAfterFailure(sheetsBefore);
      showMessageBox("数式または集計行/列を検出できませんでした。", { title: "Cht WebSheet" });
      return true;
    }
    state.hf = buildFormulaEngine(state.model);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: context.range,
      selectionAfter: context.range,
      selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
      selectionAfterRanges: activeSelectionRanges().map((item) => ({ ...item })),
      label: `${items.length} シート 自動アウトライン`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートで自動アウトラインを作成しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`自動アウトラインを作成できませんでした：${error.message || error}`);
  }
  return true;
}



function applyAutoOutlineRows(sheet, range, settings) {
  const summaryRows = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    if (outlineRowLooksLikeSummary(sheet, row, range.left, range.right)) summaryRows.push(row);
  }
  if (!summaryRows.length) return false;
  let changed = false;
  if (settings.summaryBelow) {
    let detailStart = range.top;
    summaryRows.forEach((summaryRow) => {
      const detailEnd = summaryRow - 1;
      if (detailEnd >= detailStart) changed = setOutlineLevelSpan(sheet, "rows", detailStart, detailEnd, 1) || changed;
      detailStart = summaryRow + 1;
    });
  } else {
    summaryRows.forEach((summaryRow, index) => {
      const detailStart = summaryRow + 1;
      const detailEnd = (summaryRows[index + 1] || range.bottom + 1) - 1;
      if (detailEnd >= detailStart) changed = setOutlineLevelSpan(sheet, "rows", detailStart, detailEnd, 1) || changed;
    });
  }
  return changed;
}



function applyAutoOutlineColumns(sheet, range, settings) {
  const summaryColumns = [];
  for (let col = range.left; col <= range.right; col += 1) {
    if (outlineColumnLooksLikeSummary(sheet, col, range.top, range.bottom)) summaryColumns.push(col);
  }
  if (!summaryColumns.length) return false;
  let changed = false;
  if (settings.summaryRight) {
    let detailStart = range.left;
    summaryColumns.forEach((summaryCol) => {
      const detailEnd = summaryCol - 1;
      if (detailEnd >= detailStart) changed = setOutlineLevelSpan(sheet, "columns", detailStart, detailEnd, 1) || changed;
      detailStart = summaryCol + 1;
    });
  } else {
    summaryColumns.forEach((summaryCol, index) => {
      const detailStart = summaryCol + 1;
      const detailEnd = (summaryColumns[index + 1] || range.right + 1) - 1;
      if (detailEnd >= detailStart) changed = setOutlineLevelSpan(sheet, "columns", detailStart, detailEnd, 1) || changed;
    });
  }
  return changed;
}



function outlineRowLooksLikeSummary(sheet, row, left, right) {
  let aggregateFormulaCount = 0;
  for (let col = left; col <= right; col += 1) {
    const cell = sheet.cells[cellKey(row, col)];
    if (outlineCellHasAggregateFormula(cell)) aggregateFormulaCount += 1;
    if (outlineCellHasSummaryLabel(cell)) return true;
  }
  return aggregateFormulaCount > 0;
}



function outlineColumnLooksLikeSummary(sheet, col, top, bottom) {
  let aggregateFormulaCount = 0;
  for (let row = top; row <= bottom; row += 1) {
    const cell = sheet.cells[cellKey(row, col)];
    if (outlineCellHasAggregateFormula(cell)) aggregateFormulaCount += 1;
    if (outlineCellHasSummaryLabel(cell)) return true;
  }
  return aggregateFormulaCount > 0;
}



function outlineCellHasAggregateFormula(cell) {
  const formula = String(cell?.formula || "");
  return /\b(?:SUM|SUBTOTAL|AGGREGATE|AVERAGE|COUNT|COUNTA|MAX|MIN|PRODUCT)\s*\(/i.test(formula);
}



function outlineCellHasSummaryLabel(cell) {
  const value = String(cell?.display ?? cell?.raw ?? cell?.cached ?? "");
  return /(?:合計|小計|総計|集計|total|subtotal|sum)/i.test(value);
}



function setOutlineLevelSpan(sheet, axis, start, end, level) {
  const collection = axis === "columns" ? sheet.columns : sheet.rows;
  let changed = false;
  for (let index = start; index <= end; index += 1) {
    const item = collection[index - 1] || {};
    const nextLevel = Math.max(clampOutlineLevel(item.outlineLevel), clampOutlineLevel(level));
    if (clampOutlineLevel(item.outlineLevel) !== nextLevel) changed = true;
    item.outlineLevel = nextLevel;
    collection[index - 1] = item;
  }
  return changed;
}



function clearOutlineFromSelection() {
  const sheet = activeSheet();
  if (!sheet) return;
  if (!ensureClearOutlineAllowed(sheet)) return;
  if (clearOutlineFromSelectionForSheetGroup()) return;
  const sheetBefore = captureSheetForHistory(sheet);
  let changed = false;
  changed = clearOutlineCollection(sheet.rows) || changed;
  changed = clearOutlineCollection(sheet.columns) || changed;
  if (!changed) {
    setStatus("クリアするアウトラインはありません。");
    return;
  }
  const range = { sheetIndex: state.activeSheetIndex, top: 1, bottom: Math.max(1, sheet.rowCount), left: 1, right: Math.max(1, sheet.colCount) };
  commitDataOperation(sheet, sheetBefore, range, "アウトラインのクリア", "アウトラインをクリアしました。", activeSelectionRange() || range);
}



function clearOutlineFromSelectionForSheetGroup() {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const items = sheetIndexes
    .map((sheetIndex) => ({ sheetIndex, sheet: state.model?.sheets?.[sheetIndex] }))
    .filter((item) => item.sheet);
  if (!items.length) return true;
  if (items.some(({ sheet }) => !ensureClearOutlineAllowed(sheet))) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    let changed = false;
    items.forEach(({ sheet }) => {
      changed = clearOutlineCollection(sheet.rows) || changed;
      changed = clearOutlineCollection(sheet.columns) || changed;
    });
    if (!changed) {
      setStatus("クリアするアウトラインはありません。");
      return true;
    }
    state.hf = buildFormulaEngine(state.model);
    const range = activeSelectionRange() || { sheetIndex: state.activeSheetIndex, top: 1, bottom: 1, left: 1, right: 1 };
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
      selectionAfterRanges: activeSelectionRanges().map((item) => ({ ...item })),
      label: `${items.length} シート アウトラインのクリア`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートでアウトラインをクリアしました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`アウトラインをクリアできませんでした：${error.message || error}`);
  }
  return true;
}



function clearOutlineCollection(collection) {
  let changed = false;
  collection.forEach((item) => {
    if (!item) return;
    if (clampOutlineLevel(item.outlineLevel)) {
      item.outlineLevel = 0;
      changed = true;
    }
    if (item.outlineHidden) {
      item.outlineHidden = false;
      if (item.hidden) item.hidden = false;
      changed = true;
    }
    if (item.outlineCollapsed) {
      item.outlineCollapsed = false;
      changed = true;
    }
  });
  return changed;
}



function groupSelectionOutline() {
  applyOutlineToSelection(1, "グループ化しました。");
}



function ungroupSelectionOutline() {
  applyOutlineToSelection(-1, "グループ解除しました。");
}



function applyOutlineToSelection(delta, statusMessage, axis = null, providedRange = null) {
  const context = providedRange ? { sheet: activeSheet(), range: providedRange } : selectedDataRange(delta > 0 ? "グループ化" : "グループ解除");
  if (!context) return;
  const { sheet, range } = context;
  const targetAxis = outlineAxisForRange(range, axis);
  if (!ensureOutlineAxisAllowed(sheet, targetAxis)) return;
  if (applyOutlineToSelectionForSheetGroup(context, delta, statusMessage, axis)) return;
  const sheetBefore = captureSheetForHistory(sheet);
  applyOutlineToSheet(sheet, range, delta, targetAxis === "columns");
  commitDataOperation(sheet, sheetBefore, range, `${rangeToLabel(range)} アウトライン`, statusMessage);
}



function applyOutlineToSelectionForSheetGroup(context, delta, statusMessage, axis = null) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !context?.range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const range = materializeSelectionRangeForSheet(sheet, context.range, sheetIndex);
    if (!range) continue;
    if (!ensureOutlineAxisAllowed(sheet, outlineAxisForRange(range, axis))) return true;
    items.push({ sheetIndex, sheet, range });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach(({ sheet, range }) => applyOutlineToSheet(sheet, range, delta, outlineTargetColumnsForRange(range, axis)));
    state.hf = buildFormulaEngine(state.model);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: context.range,
      selectionAfter: context.range,
      selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
      selectionAfterRanges: activeSelectionRanges().map((item) => ({ ...item })),
      label: `${items.length} シート アウトライン`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートで${statusMessage}`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`アウトラインを変更できませんでした：${error.message || error}`);
  }
  return true;
}



function outlineAxisForRange(range, axis = null) {
  if (axis === "columns" || axis === "rows") return axis;
  return range?.fullColumn
    ? "columns"
    : range?.fullRow
      ? "rows"
      : (range?.right || 1) - (range?.left || 1) > (range?.bottom || 1) - (range?.top || 1)
        ? "columns"
        : "rows";
}



function outlineAxisSheetAllows(sheet, axis) {
  return sheetProtectionAllows(sheet, axis === "columns" ? "allowFormatColumns" : "allowFormatRows");
}



function ensureOutlineAxisAllowed(sheet, axis) {
  return ensureSheetProtectionPermission(
    sheet,
    axis === "columns" ? "allowFormatColumns" : "allowFormatRows",
    axis === "columns"
      ? "保護されたシートでは列のアウトラインを変更できません。"
      : "保護されたシートでは行のアウトラインを変更できません。",
  );
}



function ensureAutoOutlineAllowed(sheet) {
  if (!isSheetProtected(sheet)) return true;
  const rowsAllowed = sheetProtectionAllows(sheet, "allowFormatRows");
  const columnsAllowed = sheetProtectionAllows(sheet, "allowFormatColumns");
  if (rowsAllowed && columnsAllowed) return true;
  showProtectedSheetMessage("保護されたシートでは自動アウトラインを作成できません。行と列の書式設定を許可してください。");
  return false;
}



function ensureClearOutlineAllowed(sheet) {
  if (!isSheetProtected(sheet)) return true;
  const rowsAllowed = sheetProtectionAllows(sheet, "allowFormatRows");
  const columnsAllowed = sheetProtectionAllows(sheet, "allowFormatColumns");
  if (rowsAllowed && columnsAllowed) return true;
  showProtectedSheetMessage("保護されたシートではアウトラインをクリアできません。行と列の書式設定を許可してください。");
  return false;
}



function outlineTargetColumnsForRange(range, axis = null) {
  return outlineAxisForRange(range, axis) === "columns";
}



function applyOutlineToSheet(sheet, range, delta, targetColumns) {
  if (targetColumns) {
    for (let col = range.left; col <= range.right; col += 1) {
      const column = sheet.columns[col - 1] || {};
      column.outlineLevel = Math.max(0, Math.min(8, Number(column.outlineLevel || 0) + delta));
      if (delta < 0 && !column.outlineLevel && column.outlineHidden) {
        column.outlineHidden = false;
        column.hidden = false;
      }
      if (delta < 0 && !column.outlineLevel && column.outlineCollapsed) {
        column.outlineCollapsed = false;
      }
      sheet.columns[col - 1] = column;
    }
  } else {
    for (let row = range.top; row <= range.bottom; row += 1) {
      const rowModel = sheet.rows[row - 1] || {};
      rowModel.outlineLevel = Math.max(0, Math.min(8, Number(rowModel.outlineLevel || 0) + delta));
      if (delta < 0 && !rowModel.outlineLevel && rowModel.outlineHidden) {
        rowModel.outlineHidden = false;
        rowModel.hidden = false;
      }
      if (delta < 0 && !rowModel.outlineLevel && rowModel.outlineCollapsed) {
        rowModel.outlineCollapsed = false;
      }
      sheet.rows[row - 1] = rowModel;
    }
  }
}



function buildOutlineControlMaps(sheet) {
  const maps = { rows: new Map(), columns: new Map(), rowLevels: [], columnLevels: [] };
  if (sheet?.showOutlineSymbols === false) return maps;
  const rowGroups = outlineGroups(sheet, "rows");
  const columnGroups = outlineGroups(sheet, "columns");
  maps.rowLevels = outlineDisplayLevelButtons(sheet, "rows", rowGroups);
  maps.columnLevels = outlineDisplayLevelButtons(sheet, "columns", columnGroups);
  rowGroups.forEach((group) => {
    const target = outlineControlIndex(sheet, group);
    if (!target) return;
    maps.rows.set(target, `${maps.rows.get(target) || ""}${outlineToggleHtml(group)}`);
  });
  columnGroups.forEach((group) => {
    const target = outlineControlIndex(sheet, group);
    if (!target) return;
    maps.columns.set(target, `${maps.columns.get(target) || ""}${outlineToggleHtml(group)}`);
  });
  return maps;
}



function outlineDisplayLevelButtons(sheet, axis, groups = outlineGroups(sheet, axis)) {
  const maxLevel = Math.max(0, ...groups.map((group) => Number(group.level) || 0));
  if (!maxLevel) return [];
  const activeLevel = outlineActiveDisplayLevel(sheet, axis, maxLevel);
  const buttons = [];
  for (let level = 1; level <= maxLevel + 1; level += 1) {
    buttons.push({ axis, level, active: level === activeLevel });
  }
  return buttons;
}



function outlineActiveDisplayLevel(sheet, axis, maxLevel = 0) {
  const collection = axis === "columns" ? sheet.columns : sheet.rows;
  const hiddenLevels = [];
  (collection || []).forEach((item) => {
    const level = clampOutlineLevel(item?.outlineLevel);
    if (level > 0 && (item?.outlineHidden === true || item?.hidden === true)) hiddenLevels.push(level);
  });
  if (!hiddenLevels.length) return Math.max(1, maxLevel + 1);
  return Math.max(1, Math.min(maxLevel + 1, Math.min(...hiddenLevels)));
}



function renderOutlineLevelButtons(buttons, className) {
  if (!buttons?.length) return "";
  const items = buttons.map((button) => outlineDisplayLevelButtonHtml(button)).join("");
  return `<div class="outline-level-buttons ${className}" aria-hidden="false">${items}</div>`;
}



function outlineDisplayLevelButtonHtml(button) {
  const axisLabel = button.axis === "columns" ? "列" : "行";
  const title = `アウトライン レベル ${button.level}（${axisLabel}）`;
  return `<button class="outline-level-button${button.active ? " is-active" : ""}" type="button" data-outline-level-command="${button.axis}" data-outline-display-level="${button.level}" title="${escapeAttr(title)}" aria-label="${escapeAttr(title)}">${button.level}</button>`;
}



function renderCornerHeader(sheet, outlineControls) {
  const rowButtons = sheet?.showOutlineSymbols === false ? "" : renderOutlineLevelButtons(outlineControls.rowLevels, "is-rows");
  const columnButtons = sheet?.showOutlineSymbols === false ? "" : renderOutlineLevelButtons(outlineControls.columnLevels, "is-columns");
  const hasControls = rowButtons || columnButtons;
  return `<div class="grid-header corner-header${hasControls ? " has-outline-level-controls" : ""}" style="grid-column:1;grid-row:1">${columnButtons}${rowButtons}</div>`;
}



function outlineGroups(sheet, axis) {
  const count = axis === "columns" ? sheet.colCount : sheet.rowCount;
  const groups = [];
  for (let level = 1; level <= 8; level += 1) {
    let index = 1;
    while (index <= count) {
      if (outlineLevelAt(sheet, axis, index) < level) {
        index += 1;
        continue;
      }
      const start = index;
      while (index <= count && outlineLevelAt(sheet, axis, index) >= level) index += 1;
      const end = index - 1;
      const collection = axis === "columns" ? sheet.columns : sheet.rows;
      const summaryIndex = outlineControlIndex(sheet, { axis, start, end });
      const summaryItem = summaryIndex ? collection[summaryIndex - 1] : null;
      groups.push({
        axis,
        level,
        start,
        end,
        collapsed: summaryItem?.outlineCollapsed === true
          || outlineGroupItems(sheet, axis, start, end).some((item) => item?.outlineHidden),
      });
    }
  }
  return groups;
}



function outlineLevelAt(sheet, axis, index) {
  const item = axis === "columns" ? sheet.columns[index - 1] : sheet.rows[index - 1];
  return Math.max(0, Math.min(8, Number(item?.outlineLevel || 0)));
}



function outlineGroupItems(sheet, axis, start, end) {
  const collection = axis === "columns" ? sheet.columns : sheet.rows;
  const items = [];
  for (let index = start; index <= end; index += 1) {
    items.push(collection[index - 1] || {});
  }
  return items;
}



function outlineControlIndex(sheet, group) {
  const count = group.axis === "columns" ? sheet.colCount : sheet.rowCount;
  const settings = normalizeOutlineSettings(sheet.outlineSettings);
  const summaryAfter = group.axis === "columns" ? settings.summaryRight : settings.summaryBelow;
  if (summaryAfter) return group.end < count ? group.end + 1 : group.end;
  return group.start > 1 ? group.start - 1 : group.start;
}



function outlineToggleHtml(group) {
  const label = group.axis === "columns"
    ? `${columnName(group.start)}:${columnName(group.end)}`
    : `${group.start}:${group.end}`;
  const symbol = group.collapsed ? "+" : "−";
  const title = group.collapsed ? "詳細を表示" : "詳細を非表示";
  return `<button class="outline-toggle is-level-${group.level}${group.collapsed ? " is-collapsed" : ""}" type="button" data-outline-toggle="${group.axis}" data-outline-start="${group.start}" data-outline-end="${group.end}" data-outline-level="${group.level}" data-outline-collapsed="${group.collapsed ? "true" : "false"}" title="${escapeAttr(`${title}: ${label}`)}" aria-label="${escapeAttr(`${title}: ${label}`)}">${symbol}</button>`;
}



function handleOutlineToggleClick(event) {
  event.preventDefault();
  event.stopPropagation();
  const button = event.currentTarget;
  const axis = button.dataset.outlineToggle;
  const start = Number(button.dataset.outlineStart);
  const end = Number(button.dataset.outlineEnd);
  const level = Number(button.dataset.outlineLevel || 1);
  const collapse = button.dataset.outlineCollapsed !== "true";
  if (!axis || !Number.isFinite(start) || !Number.isFinite(end)) return;
  setOutlineGroupCollapsed(axis, start, end, level, collapse);
}



function handleOutlineLevelCommandClick(event) {
  event.preventDefault();
  event.stopPropagation();
  const button = event.currentTarget;
  const axis = button.dataset.outlineLevelCommand;
  const level = Math.max(1, Math.round(Number(button.dataset.outlineDisplayLevel) || 1));
  if (axis !== "rows" && axis !== "columns") return;
  setOutlineDisplayLevel(axis, level);
}



function setOutlineDisplayLevel(axis, level) {
  const sheet = activeSheet();
  if (!sheet) return;
  if (!ensureOutlineAxisAllowed(sheet, axis)) return;
  if (setOutlineDisplayLevelForSheetGroup(axis, level)) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const changed = setOutlineDisplayLevelOnSheet(sheet, axis, level);
  if (!changed) {
    setStatus(`アウトラインは既にレベル ${level} で表示されています。`);
    return;
  }
  const range = outlineRangeForAxis(axis, 1, axis === "columns" ? sheet.colCount : sheet.rowCount);
  commitDataOperation(
    sheet,
    sheetBefore,
    range,
    `${axis === "columns" ? "列" : "行"}アウトライン レベル ${level}`,
    `アウトラインをレベル ${level} で表示しました。`,
    range,
  );
}



function setOutlineDisplayLevelOnSheet(sheet, axis, level) {
  const targetLevel = Math.max(1, Math.round(Number(level) || 1));
  const groups = outlineGroups(sheet, axis);
  if (!groups.length) return false;
  let changed = false;
  groups
    .filter((group) => group.level < targetLevel)
    .sort((left, right) => left.level - right.level || left.start - right.start)
    .forEach((group) => {
      changed = setOutlineGroupCollapsedOnSheet(sheet, axis, group.start, group.end, group.level, false) || changed;
    });
  groups
    .filter((group) => group.level >= targetLevel)
    .sort((left, right) => right.level - left.level || left.start - right.start)
    .forEach((group) => {
      changed = setOutlineGroupCollapsedOnSheet(sheet, axis, group.start, group.end, group.level, true) || changed;
    });
  return changed;
}



function setOutlineDisplayLevelForSheetGroup(axis, level) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const items = sheetIndexes
    .map((sheetIndex) => ({ sheetIndex, sheet: state.model?.sheets?.[sheetIndex] }))
    .filter((item) => item.sheet);
  if (!items.length) return true;
  if (items.some(({ sheet }) => !ensureOutlineAxisAllowed(sheet, axis))) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    let changed = false;
    items.forEach(({ sheet }) => {
      changed = setOutlineDisplayLevelOnSheet(sheet, axis, level) || changed;
    });
    if (!changed) {
      setStatus(`アウトラインは既にレベル ${level} で表示されています。`);
      return true;
    }
    state.hf = buildFormulaEngine(state.model);
    const sheet = activeSheet();
    const range = outlineRangeForAxis(axis, 1, axis === "columns" ? sheet.colCount : sheet.rowCount);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
      selectionAfterRanges: activeSelectionRanges().map((item) => ({ ...item })),
      label: `${items.length} シート ${axis === "columns" ? "列" : "行"}アウトライン レベル ${level}`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートでアウトラインをレベル ${level} で表示しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`アウトライン レベルを変更できませんでした：${error.message || error}`);
  }
  return true;
}



function setOutlineGroupCollapsed(axis, start, end, level, collapse) {
  const sheet = activeSheet();
  if (!sheet) return;
  if (!ensureOutlineAxisAllowed(sheet, axis)) return;
  if (setOutlineGroupCollapsedForSheetGroup(axis, start, end, level, collapse)) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const changed = setOutlineGroupCollapsedOnSheet(sheet, axis, start, end, level, collapse);
  if (!changed) {
    setStatus(collapse ? "非表示にする詳細はありません。" : "表示する詳細はありません。");
    return;
  }
  const range = outlineRangeForAxis(axis, start, end);
  commitDataOperation(
    sheet,
    sheetBefore,
    range,
    `${rangeToLabel(range)} ${collapse ? "詳細非表示" : "詳細表示"}`,
    collapse ? "アウトラインの詳細を非表示にしました。" : "アウトラインの詳細を表示しました。",
    range,
  );
}



function setOutlineGroupCollapsedOnSheet(sheet, axis, start, end, level, collapse) {
  const collection = axis === "columns" ? sheet.columns : sheet.rows;
  let changed = false;
  for (let index = start; index <= end; index += 1) {
    const item = collection[index - 1] || {};
    if (outlineLevelAt(sheet, axis, index) < level) continue;
    if (collapse) {
      if (!item.outlineHidden || !item.hidden) changed = true;
      storeOutlineVisibleSize(sheet, axis, item);
      item.outlineHidden = true;
      item.hidden = true;
    } else if (item.outlineHidden) {
      item.outlineHidden = false;
      item.hidden = false;
      restoreOutlineVisibleSize(sheet, axis, item);
      changed = true;
    }
    collection[index - 1] = item;
  }
  const summaryIndex = outlineControlIndex(sheet, { axis, start, end });
  if (summaryIndex) {
    const summary = collection[summaryIndex - 1] || {};
    const nextCollapsed = collapse === true;
    if (summary.outlineCollapsed !== nextCollapsed) changed = true;
    if (nextCollapsed) {
      summary.outlineCollapsed = true;
    } else if (summary.outlineCollapsed) {
      summary.outlineCollapsed = false;
    }
    collection[summaryIndex - 1] = summary;
  }
  return changed;
}



function storeOutlineVisibleSize(sheet, axis, item) {
  if (!item) return;
  if (axis === "columns") {
    const width = Number(item.width);
    if (Number.isFinite(width) && width > 0) item.outlineRestoreWidth = width;
  } else {
    const height = Number(item.height);
    if (Number.isFinite(height) && height > 0) item.outlineRestoreHeight = height;
  }
}



function restoreOutlineVisibleSize(sheet, axis, item) {
  if (!item) return;
  if (axis === "columns") {
    const width = Number(item.width);
    if (Number.isFinite(width) && width > 0) return;
    const restored = Number(item.outlineRestoreWidth) || Number(sheet?.defaultColumnWidth) || columnWidthToPixels(8.43, sheet?.displayScale || 1, sheet?.widthMetrics?.maxDigitWidth || 7);
    item.width = Math.max(1, restored);
    return;
  }
  const height = Number(item.height);
  if (Number.isFinite(height) && height > 0) return;
  const restored = Number(item.outlineRestoreHeight) || Number(sheet?.defaultRowHeight) || rowHeightToPixels(15, sheet?.displayScale || 1);
  item.height = Math.max(1, restored);
}



function setOutlineGroupCollapsedForSheetGroup(axis, start, end, level, collapse) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const items = sheetIndexes
    .map((sheetIndex) => ({ sheetIndex, sheet: state.model?.sheets?.[sheetIndex] }))
    .filter((item) => item.sheet);
  if (!items.length) return true;
  if (items.some(({ sheet }) => !ensureOutlineAxisAllowed(sheet, axis))) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    let changed = false;
    items.forEach(({ sheet }) => {
      changed = setOutlineGroupCollapsedOnSheet(sheet, axis, start, end, level, collapse) || changed;
    });
    if (!changed) {
      setStatus(collapse ? "非表示にする詳細はありません。" : "表示する詳細はありません。");
      return true;
    }
    state.hf = buildFormulaEngine(state.model);
    const range = outlineRangeForAxis(axis, start, end);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
      selectionAfterRanges: activeSelectionRanges().map((item) => ({ ...item })),
      label: `${items.length} シート ${collapse ? "詳細非表示" : "詳細表示"}`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートでアウトラインの詳細を${collapse ? "非表示" : "表示"}にしました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`アウトラインの詳細を変更できませんでした：${error.message || error}`);
  }
  return true;
}

export {
  addForecastChartToSheet,
  applyAutoOutlineColumns,
  applyAutoOutlineRows,
  applyForecastSheetDialog,
  applyOutlineAutomaticStyles,
  applyOutlineDialog,
  applyOutlineSettingsDialog,
  applyOutlineSettingsForSheetGroup,
  applyOutlineToSelection,
  applyOutlineToSelectionForSheetGroup,
  applyOutlineToSheet,
  autoOutlineSelection,
  autoOutlineSelectionForSheetGroup,
  buildOutlineControlMaps,
  clearOutlineCollection,
  clearOutlineFromSelection,
  clearOutlineFromSelectionForSheetGroup,
  createForecastSheetFromSelection,
  dataAnalysisAnovaRows,
  dataAnalysisAnovaSummaryRow,
  dataAnalysisAnovaTwoFactorNoReplicationRows,
  dataAnalysisAnovaTwoFactorReplicationRows,
  dataAnalysisBetaFraction,
  dataAnalysisCellNumericValue,
  dataAnalysisChartFromOutputRows,
  dataAnalysisCleanSignedZero,
  dataAnalysisComplexText,
  dataAnalysisComplexValueFromCell,
  dataAnalysisConfidenceInterval,
  dataAnalysisConfidencePercentLabel,
  dataAnalysisCorrelation,
  dataAnalysisCovarianceP,
  dataAnalysisCumulativeStandardError,
  dataAnalysisDescriptiveRows,
  dataAnalysisDiscreteDistribution,
  dataAnalysisErf,
  dataAnalysisExponentialSmoothingRows,
  dataAnalysisFInverseRightTail,
  dataAnalysisFRightTail,
  dataAnalysisFTestRows,
  dataAnalysisFourierInputValues,
  dataAnalysisFourierRows,
  dataAnalysisHistogramRows,
  dataAnalysisInputMatrix,
  dataAnalysisInverseNormal,
  dataAnalysisIsPowerOfTwo,
  dataAnalysisKthValue,
  dataAnalysisKurtosis,
  dataAnalysisLogGamma,
  dataAnalysisMatrixInverse,
  dataAnalysisMatrixMultiply,
  dataAnalysisMatrixRows,
  dataAnalysisMatrixVectorMultiply,
  dataAnalysisMean,
  dataAnalysisMedian,
  dataAnalysisMode,
  dataAnalysisMovingAverageRows,
  dataAnalysisNormalCdf,
  dataAnalysisNormalizeComplexValue,
  dataAnalysisOls,
  dataAnalysisPairedNumericValues,
  dataAnalysisParseComplexValue,
  dataAnalysisPatternValues,
  dataAnalysisRandomDiscrete,
  dataAnalysisRandomDistributionLabel,
  dataAnalysisRandomGenerator,
  dataAnalysisRandomNormal,
  dataAnalysisRandomPoisson,
  dataAnalysisRandomRows,
  dataAnalysisRandomValue,
  dataAnalysisRankPercentileRows,
  dataAnalysisRegressionObservations,
  dataAnalysisRegressionRows,
  dataAnalysisRegularizedBeta,
  dataAnalysisRound,
  dataAnalysisSamplingRows,
  dataAnalysisSeries,
  dataAnalysisSeriesValueMap,
  dataAnalysisSkewness,
  dataAnalysisStdevS,
  dataAnalysisStudentTInverseTwoTail,
  dataAnalysisStudentTTwoTail,
  dataAnalysisSum,
  dataAnalysisTTestRows,
  dataAnalysisTTestTitle,
  dataAnalysisTranspose,
  dataAnalysisVarianceP,
  dataAnalysisVarianceS,
  dataAnalysisZTestRows,
  ensureAutoOutlineAllowed,
  ensureClearOutlineAllowed,
  ensureOutlineAxisAllowed,
  forecastAggregateDuplicateValues,
  forecastAggregateLabel,
  forecastConfidenceSpread,
  forecastDialogDefaults,
  forecastDialogOptions,
  forecastDialogPreview,
  forecastDuplicateTimelineCount,
  forecastEndValueFromDialogText,
  forecastNumericValue,
  forecastPeriodsForOptions,
  forecastResidualStandardDeviation,
  forecastRowsForOptions,
  forecastSeriesPointsForRange,
  forecastStep,
  forecastTimelineOutputLabel,
  forecastTimelineOutputMetadata,
  forecastTimelineOutputValue,
  forecastTimelineValue,
  groupSelectionOutline,
  handleForecastDialogSync,
  handleOutlineDialogSync,
  handleOutlineLevelCommandClick,
  handleOutlineToggleClick,
  interpolateForecastMissingPoints,
  linearForecast,
  linearRegressionForForecast,
  normalizeForecastPoints,
  openForecastSheetDialog,
  openOutlineGroupDialog,
  openOutlineSettingsDialog,
  outlineActiveDisplayLevel,
  outlineAxisForRange,
  outlineAxisSheetAllows,
  outlineCellHasAggregateFormula,
  outlineCellHasSummaryLabel,
  outlineColumnLooksLikeSummary,
  outlineControlIndex,
  outlineDialogValidation,
  outlineDisplayLevelButtonHtml,
  outlineDisplayLevelButtons,
  outlineGroupItems,
  outlineGroups,
  outlineLevelAt,
  outlineRowLooksLikeSummary,
  outlineTargetColumnsForRange,
  outlineToggleHtml,
  renderCornerHeader,
  renderOutlineLevelButtons,
  restoreOutlineVisibleSize,
  restoreScenarioSummarySnapshots,
  scenarioPointAddressKey,
  scenarioPointDisplay,
  scenarioPointKey,
  scenarioPointLabel,
  scenarioPointValueKey,
  scenarioSummaryAddress,
  scenarioSummaryValueMaps,
  scenarioValueForPoint,
  seasonalLinearForecast,
  setOutlineDisplayLevel,
  setOutlineDisplayLevelForSheetGroup,
  setOutlineDisplayLevelOnSheet,
  setOutlineGroupCollapsed,
  setOutlineGroupCollapsedForSheetGroup,
  setOutlineGroupCollapsedOnSheet,
  setOutlineLevelSpan,
  setScenarioValueForPoint,
  storeOutlineVisibleSize,
  syncForecastDialog,
  syncOutlineDialog,
  syncOutlineSettingsDialog,
  ungroupSelectionOutline,
  workbookHasForecastSheet,
};
