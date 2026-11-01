// Feature module; shared bindings are imported explicitly.
import {
  ArraySize,
  CellError,
  ErrorType,
  FunctionPlugin,
  SimpleRangeValue,
  WEBSHEET_CONVERT_UNITS,
  WEBSHEET_FORMULA_NA_SENTINEL,
  WEBSHEET_FORMULA_RANGE,
  WebSheetFormulaPlugin,
  state,
} from "../../app-core.js";
import {
  dataAnalysisComplexText,
  dataAnalysisStudentTInverseTwoTail,
} from "../data/forecast-and-statistics.js";
import {
  decodeFormulaRangeForSheet,
  formulaArgumentBoolean,
  formulaArgumentNumber,
  formulaRangeValueMatrix,
  formulaScalarValueForCell,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  textPartFormulaValue,
  textSplitFormulaMatrix,
  webSheetFormulaAValues,
  webSheetFormulaAccrint,
  webSheetFormulaAccrintM,
  webSheetFormulaAcot,
  webSheetFormulaAddMonths,
  webSheetFormulaAggregate,
  webSheetFormulaAmorDegrc,
  webSheetFormulaAmorLinc,
  webSheetFormulaApplyLambda,
  webSheetFormulaApproximateLookup,
  webSheetFormulaArabic,
  webSheetFormulaArraySizeFromFirstArg,
  webSheetFormulaAstBooleanLiteral,
  webSheetFormulaAstNumberLiteral,
  webSheetFormulaAstRuntimeMatrix,
  webSheetFormulaAstRuntimeScalar,
  webSheetFormulaAstStaticMatrix,
  webSheetFormulaAstStringLiteral,
  webSheetFormulaAverageDeviation,
  webSheetFormulaAverageIf,
  webSheetFormulaBase,
  webSheetFormulaBesselI,
  webSheetFormulaBesselJ,
  webSheetFormulaBesselK,
  webSheetFormulaBesselY,
  webSheetFormulaBetaDist,
  webSheetFormulaBetaInv,
  webSheetFormulaBinomDist,
  webSheetFormulaBinomDistRange,
  webSheetFormulaBinomInv,
  webSheetFormulaBitShift,
  webSheetFormulaBitwise,
  webSheetFormulaByColArraySize,
  webSheetFormulaByRowArraySize,
  webSheetFormulaCeilingMath,
  webSheetFormulaCeilingPrecise,
  webSheetFormulaCellInfoFormat,
  webSheetFormulaCellInfoType,
  webSheetFormulaChiSqDist,
  webSheetFormulaChiSqInv,
  webSheetFormulaChiSqTest,
  webSheetFormulaCombinationA,
  webSheetFormulaCombinationResult,
  webSheetFormulaCompareForSort,
  webSheetFormulaCompareValues,
  webSheetFormulaComplex,
  webSheetFormulaComplexDivideValues,
  webSheetFormulaComplexText,
  webSheetFormulaComplexTextOrError,
  webSheetFormulaComplexUnary,
  webSheetFormulaConditionalValues,
  webSheetFormulaConfidenceNorm,
  webSheetFormulaConfidenceT,
  webSheetFormulaConvertBase,
  webSheetFormulaConvertTemperature,
  webSheetFormulaConvertUnitKey,
  webSheetFormulaCorrelation,
  webSheetFormulaCountBlank,
  webSheetFormulaCoupDayBs,
  webSheetFormulaCoupDays,
  webSheetFormulaCoupDaysNc,
  webSheetFormulaCoupNcd,
  webSheetFormulaCoupNum,
  webSheetFormulaCoupPcd,
  webSheetFormulaCovariance,
  webSheetFormulaDatabaseAggregate,
  webSheetFormulaDateParts,
  webSheetFormulaDateSerial,
  webSheetFormulaDateValue,
  webSheetFormulaDays360,
  webSheetFormulaDb,
  webSheetFormulaDdb,
  webSheetFormulaDecimal,
  webSheetFormulaDecimalToSignedBase,
  webSheetFormulaDevSq,
  webSheetFormulaDisc,
  webSheetFormulaDollarDe,
  webSheetFormulaDollarFr,
  webSheetFormulaDoubleFactorial,
  webSheetFormulaErf,
  webSheetFormulaError,
  webSheetFormulaExponDist,
  webSheetFormulaFDist,
  webSheetFormulaFInv,
  webSheetFormulaFTest,
  webSheetFormulaFactorial,
  webSheetFormulaFactorialResult,
  webSheetFormulaFlattenValues,
  webSheetFormulaFloorMath,
  webSheetFormulaFloorPrecise,
  webSheetFormulaFormatFixed,
  webSheetFormulaFrequency,
  webSheetFormulaFutureValue,
  webSheetFormulaFvSchedule,
  webSheetFormulaGamma,
  webSheetFormulaGammaDist,
  webSheetFormulaGammaInv,
  webSheetFormulaGeoMean,
  webSheetFormulaGreatestCommonDivisor,
  webSheetFormulaHarMean,
  webSheetFormulaHypergeomDist,
  webSheetFormulaImabs,
  webSheetFormulaImproduct,
  webSheetFormulaImsum,
  webSheetFormulaInfoN,
  webSheetFormulaIntrate,
  webSheetFormulaInverseStandardNormal,
  webSheetFormulaIpmt,
  webSheetFormulaIrr,
  webSheetFormulaIsEmptyValue,
  webSheetFormulaIsoWeeknumValue,
  webSheetFormulaIspmt,
  webSheetFormulaKurt,
  webSheetFormulaLeastCommonMultiple,
  webSheetFormulaLetAstValue,
  webSheetFormulaLetName,
  webSheetFormulaLetScalarNumber,
  webSheetFormulaLogGammaPositive,
  webSheetFormulaLogNormDist,
  webSheetFormulaLogNormInv,
  webSheetFormulaLookupEntries,
  webSheetFormulaLookupMatchIndex,
  webSheetFormulaMakeArraySize,
  webSheetFormulaMatrix,
  webSheetFormulaMatrixDeterminant,
  webSheetFormulaMatrixDimensions,
  webSheetFormulaMatrixIdentity,
  webSheetFormulaMatrixInverse,
  webSheetFormulaMatrixMultiply,
  webSheetFormulaMirr,
  webSheetFormulaMod,
  webSheetFormulaModeMultValues,
  webSheetFormulaModeSnglValue,
  webSheetFormulaMround,
  webSheetFormulaMultipleRegression,
  webSheetFormulaNegbinomDist,
  webSheetFormulaNetworkDays,
  webSheetFormulaNormalCdf,
  webSheetFormulaNormalInv,
  webSheetFormulaNormalPdf,
  webSheetFormulaNper,
  webSheetFormulaNpv,
  webSheetFormulaNumber,
  webSheetFormulaNumericValues,
  webSheetFormulaPayment,
  webSheetFormulaPduration,
  webSheetFormulaPercentRank,
  webSheetFormulaPercentile,
  webSheetFormulaPermutation,
  webSheetFormulaPermutationA,
  webSheetFormulaPoissonDist,
  webSheetFormulaPpmt,
  webSheetFormulaPresentValue,
  webSheetFormulaPriceDisc,
  webSheetFormulaPriceMat,
  webSheetFormulaProb,
  webSheetFormulaRank,
  webSheetFormulaRate,
  webSheetFormulaRawValue,
  webSheetFormulaReceived,
  webSheetFormulaReferenceCell,
  webSheetFormulaReferenceFromAstArg,
  webSheetFormulaRegression,
  webSheetFormulaRegressionFeatureCountFromSizes,
  webSheetFormulaRegressionPredict,
  webSheetFormulaRegressionResultMatrix,
  webSheetFormulaRoman,
  webSheetFormulaRoundAwayToParity,
  webSheetFormulaRoundToDecimals,
  webSheetFormulaRowKey,
  webSheetFormulaRri,
  webSheetFormulaScalar,
  webSheetFormulaSecurityDuration,
  webSheetFormulaSecurityPrice,
  webSheetFormulaSecurityYield,
  webSheetFormulaSeriesSum,
  webSheetFormulaSheetIndexFromFormulaSheetId,
  webSheetFormulaSkew,
  webSheetFormulaSln,
  webSheetFormulaStdevFromArgs,
  webSheetFormulaSteyx,
  webSheetFormulaStudentTCdf,
  webSheetFormulaStudentTDistTwoTail,
  webSheetFormulaStudentTInv,
  webSheetFormulaStudentTPdf,
  webSheetFormulaSumIf,
  webSheetFormulaSumPairwiseSquares,
  webSheetFormulaSumProduct,
  webSheetFormulaSyd,
  webSheetFormulaTTest,
  webSheetFormulaTbillEq,
  webSheetFormulaTbillPrice,
  webSheetFormulaTbillYield,
  webSheetFormulaText,
  webSheetFormulaTimeUnit,
  webSheetFormulaTimeValue,
  webSheetFormulaTrigReciprocal,
  webSheetFormulaTrimMean,
  webSheetFormulaVarianceFromArgs,
  webSheetFormulaVdb,
  webSheetFormulaWeekdayValue,
  webSheetFormulaWeeknumValue,
  webSheetFormulaWeibullDist,
  webSheetFormulaWorkbookFileName,
  webSheetFormulaWorkday,
  webSheetFormulaXirr,
  webSheetFormulaXlookupResult,
  webSheetFormulaXnpv,
  webSheetFormulaYearFrac,
  webSheetFormulaYieldDisc,
  webSheetFormulaYieldMat,
  webSheetFormulaZTest,
} from "./calculations.js";
import {
  cellHasFormulaContent,
} from "../ui/annotations-and-cell-format.js";
import {
  columnName,
  columnWidthToPixels,
} from "../ui/help-and-utilities.js";
import {
  chooseColsFormulaMatrix,
  chooseRowsFormulaMatrix,
  dropFormulaMatrix,
  excelDatedif,
  expandFormulaMatrix,
  filterFormulaMatrix,
  getSheetId,
  hstackFormulaMatrices,
  sequenceFormulaMatrix,
  sortByFormulaMatrix,
  takeFormulaMatrix,
  tocolFormulaMatrix,
  torowFormulaMatrix,
  vstackFormulaMatrices,
  wrapColsFormulaMatrix,
  wrapRowsFormulaMatrix,
} from "../workbook/saving-and-export.js";
function createWebSheetFormulaPluginClass() {
  return class WebSheetFormulaPlugin extends FunctionPlugin {
  textjoin(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TEXTJOIN"), (delimiter, ignoreEmpty, ...items) => {
      const parts = [];
      for (const item of items) {
        for (const value of webSheetFormulaFlattenValues(item)) {
          const text = webSheetFormulaText(value);
          if (text instanceof CellError) return text;
          if (ignoreEmpty && text === "") continue;
          parts.push(text);
        }
      }
      return parts.join(String(delimiter ?? ""));
    });
  }

  concat(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CONCAT"), (...items) => {
      let output = "";
      for (const item of items) {
        for (const value of webSheetFormulaFlattenValues(item)) {
          const text = webSheetFormulaText(value);
          if (text instanceof CellError) return text;
          output += text;
        }
      }
      return output;
    });
  }

  concatenate(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CONCATENATE"), (...items) => {
      let output = "";
      for (const item of items) {
        for (const value of webSheetFormulaFlattenValues(item)) {
          const text = webSheetFormulaText(value);
          if (text instanceof CellError) return text;
          output += text;
        }
      }
      return output;
    });
  }

  replaceText(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("REPLACE"), (oldText, startNum, numChars, newText) => {
      const source = String(oldText ?? "");
      const start = Math.trunc(Number(startNum));
      const count = Math.trunc(Number(numChars));
      if (!Number.isFinite(start) || !Number.isFinite(count) || start < 1 || count < 0) return webSheetFormulaError(ErrorType.VALUE);
      return `${source.slice(0, start - 1)}${String(newText ?? "")}${source.slice(start - 1 + count)}`;
    });
  }

  proper(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PROPER"), (text) =>
      String(text ?? "")
        .toLocaleLowerCase()
        .replace(/(^|[^0-9\p{L}])(\p{L})/gu, (match, prefix, letter) => `${prefix}${letter.toLocaleUpperCase()}`),
    );
  }

  exact(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("EXACT"), (left, right) => String(left ?? "") === String(right ?? ""));
  }

  char(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CHAR"), (number) => {
      const code = Math.trunc(Number(number));
      return code >= 1 && code <= 255 ? String.fromCharCode(code) : webSheetFormulaError(ErrorType.VALUE);
    });
  }

  code(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CODE"), (text) => {
      const value = String(text ?? "");
      return value ? value.charCodeAt(0) : webSheetFormulaError(ErrorType.VALUE);
    });
  }

  clean(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CLEAN"), (text) => String(text ?? "").replace(/[\x00-\x1F]/g, ""));
  }

  rept(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("REPT"), (text, numberTimes) => {
      const count = Math.trunc(Number(numberTimes));
      if (!Number.isFinite(count) || count < 0) return webSheetFormulaError(ErrorType.VALUE);
      const output = String(text ?? "").repeat(count);
      return output.length > 32767 ? webSheetFormulaError(ErrorType.VALUE) : output;
    });
  }

  textsplit(ast, state) {
    const text = webSheetFormulaAstRuntimeScalar(ast.args?.[0], state, this, "");
    const colDelimiter = webSheetFormulaAstRuntimeMatrix(ast.args?.[1], state, this, [[""]]);
    const rowDelimiter = ast.args?.[2]?.type === "EMPTY" ? undefined : webSheetFormulaAstRuntimeMatrix(ast.args?.[2], state, this, undefined);
    const ignoreEmpty = formulaArgumentBoolean(webSheetFormulaAstRuntimeScalar(ast.args?.[3], state, this, false), false);
    const matchMode = Math.trunc(formulaArgumentNumber(webSheetFormulaAstRuntimeScalar(ast.args?.[4], state, this, 0), 0));
    const padWith = ast.args?.[5]?.type === "EMPTY" ? WEBSHEET_FORMULA_NA_SENTINEL : webSheetFormulaAstRuntimeScalar(ast.args?.[5], state, this, WEBSHEET_FORMULA_NA_SENTINEL);
    return SimpleRangeValue.onlyValues(textSplitFormulaMatrix(text, colDelimiter, rowDelimiter, ignoreEmpty, matchMode, padWith));
  }

  textsplitArraySize(ast, state) {
    const text = webSheetFormulaAstStringLiteral(ast.args?.[0]);
    const colDelimiter = webSheetFormulaAstStaticMatrix(ast.args?.[1]) || [[webSheetFormulaAstStringLiteral(ast.args?.[1])]];
    const rowDelimiter = ast.args?.[2]?.type === "EMPTY" ? undefined : (webSheetFormulaAstStaticMatrix(ast.args?.[2]) || [[webSheetFormulaAstStringLiteral(ast.args?.[2])]]);
    const ignoreEmpty = webSheetFormulaAstBooleanLiteral(ast.args?.[3], false);
    if (!text || !colDelimiter) return new ArraySize(1, 1);
    const matrix = textSplitFormulaMatrix(text, colDelimiter, rowDelimiter, ignoreEmpty, webSheetFormulaAstNumberLiteral(ast.args?.[4], 0), "");
    const { height, width } = webSheetFormulaMatrixDimensions(matrix);
    return new ArraySize(Math.max(1, width), Math.max(1, height));
  }

  textbefore(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TEXTBEFORE"), (text, delimiter, instanceNum, matchMode, matchEnd, ifNotFound) =>
      textPartFormulaValue(text, delimiter, instanceNum, matchMode, matchEnd, ifNotFound, false),
    );
  }

  textafter(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TEXTAFTER"), (text, delimiter, instanceNum, matchMode, matchEnd, ifNotFound) =>
      textPartFormulaValue(text, delimiter, instanceNum, matchMode, matchEnd, ifNotFound, true),
    );
  }

  aggregate(ast, state) {
    if ((ast.args || []).length < 3) return webSheetFormulaError(ErrorType.NA);
    const evaluatedArgs = ast.args.map((arg) => this.evaluateAst(arg, state));
    const scalarValue = (value) => value instanceof SimpleRangeValue ? value.valuesFromTopLeftCorner()[0] : value;
    const functionNum = webSheetFormulaNumber(scalarValue(evaluatedArgs[0]));
    if (functionNum instanceof CellError) return functionNum;
    const options = webSheetFormulaNumber(scalarValue(evaluatedArgs[1]));
    if (options instanceof CellError) return options;
    return webSheetFormulaAggregate(functionNum, options, evaluatedArgs.slice(2));
  }

  ifs(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IFS"), (...args) => {
      for (let index = 0; index < args.length; index += 2) {
        if (args[index]) return args[index + 1];
      }
      return webSheetFormulaError(ErrorType.NA);
    });
  }

  switchFunction(ast, state) {
    const args = ast.args || [];
    if (args.length < 3) return webSheetFormulaError(ErrorType.NA);
    const expression = this.evaluateAst(args[0], state);
    const hasDefault = (args.length - 1) % 2 === 1;
    const pairEnd = hasDefault ? args.length - 1 : args.length;
    for (let index = 1; index < pairEnd; index += 2) {
      const candidate = this.evaluateAst(args[index], state);
      if (webSheetFormulaCompareValues(expression, candidate) === 0) return this.evaluateAst(args[index + 1], state);
    }
    return hasDefault ? this.evaluateAst(args[args.length - 1], state) : webSheetFormulaError(ErrorType.NA);
  }

  xor(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("XOR"), (...args) =>
      args.filter((value) => Boolean(value)).length % 2 === 1,
    );
  }

  letFunction(ast, state) {
    const args = ast.args || [];
    if (args.length < 3 || args.length % 2 === 0) return webSheetFormulaError(ErrorType.VALUE);
    const env = new Map();
    for (let index = 0; index < args.length - 1; index += 2) {
      const name = webSheetFormulaLetName(args[index]);
      if (!name) return webSheetFormulaError(ErrorType.VALUE);
      const value = webSheetFormulaLetAstValue(args[index + 1], state, this, env);
      if (value instanceof CellError) return value;
      env.set(name, value);
    }
    return webSheetFormulaLetAstValue(args[args.length - 1], state, this, env);
  }

  lambda() {
    return webSheetFormulaError(ErrorType.VALUE);
  }

  sumproduct(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SUMPRODUCT"), (...args) => webSheetFormulaSumProduct(args));
  }

  countblank(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COUNTBLANK"), (range) => webSheetFormulaCountBlank(range));
  }

  countif(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COUNTIF"), (range, criteria) => {
      const values = webSheetFormulaConditionalValues(range, [range, criteria], { numericOnly: false });
      return values instanceof CellError ? values : values.length;
    });
  }

  countifs(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COUNTIFS"), (...criteriaArgs) => {
      const firstRange = criteriaArgs[0];
      const values = webSheetFormulaConditionalValues(firstRange, criteriaArgs, { numericOnly: false });
      return values instanceof CellError ? values : values.length;
    });
  }

  sumif(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SUMIF"), (range, criteria, sumRange) =>
      webSheetFormulaSumIf(range, criteria, sumRange === WEBSHEET_FORMULA_NA_SENTINEL ? range : sumRange),
    );
  }

  sumifs(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SUMIFS"), (sumRange, ...criteriaArgs) => {
      const values = webSheetFormulaConditionalValues(sumRange, criteriaArgs);
      return values instanceof CellError ? values : values.reduce((sum, value) => sum + value, 0);
    });
  }

  averageif(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("AVERAGEIF"), (range, criteria, averageRange) =>
      webSheetFormulaAverageIf(range, criteria, averageRange === WEBSHEET_FORMULA_NA_SENTINEL ? range : averageRange),
    );
  }

  averageifs(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("AVERAGEIFS"), (averageRange, ...criteriaArgs) => {
      const matched = webSheetFormulaConditionalValues(averageRange, criteriaArgs);
      if (matched instanceof CellError) return matched;
      if (!matched.length) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
      return matched.reduce((sum, value) => sum + value, 0) / matched.length;
    });
  }

  averagea(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("AVERAGEA"), (...args) => {
      const values = webSheetFormulaAValues(args);
      if (values instanceof CellError) return values;
      return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : webSheetFormulaError(ErrorType.DIV_BY_ZERO);
    });
  }

  maxa(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MAXA"), (...args) => {
      const values = webSheetFormulaAValues(args);
      return values instanceof CellError ? values : values.length ? Math.max(...values) : 0;
    });
  }

  mina(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MINA"), (...args) => {
      const values = webSheetFormulaAValues(args);
      return values instanceof CellError ? values : values.length ? Math.min(...values) : 0;
    });
  }

  stdevS(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("STDEV.S"), (...args) => webSheetFormulaStdevFromArgs(args, true));
  }

  stdevP(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("STDEV.P"), (...args) => webSheetFormulaStdevFromArgs(args, false));
  }

  stdev(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("STDEV"), (...args) => webSheetFormulaStdevFromArgs(args, true));
  }

  stdevp(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("STDEVP"), (...args) => webSheetFormulaStdevFromArgs(args, false));
  }

  stdeva(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("STDEVA"), (...args) => webSheetFormulaStdevFromArgs(args, true, true));
  }

  stdevpa(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("STDEVPA"), (...args) => webSheetFormulaStdevFromArgs(args, false, true));
  }

  varS(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("VAR.S"), (...args) => webSheetFormulaVarianceFromArgs(args, true));
  }

  varP(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("VAR.P"), (...args) => webSheetFormulaVarianceFromArgs(args, false));
  }

  varLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("VAR"), (...args) => webSheetFormulaVarianceFromArgs(args, true));
  }

  varp(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("VARP"), (...args) => webSheetFormulaVarianceFromArgs(args, false));
  }

  vara(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("VARA"), (...args) => webSheetFormulaVarianceFromArgs(args, true, true));
  }

  varpa(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("VARPA"), (...args) => webSheetFormulaVarianceFromArgs(args, false, true));
  }

  correl(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CORREL"), (array1, array2) => webSheetFormulaCorrelation(array1, array2));
  }

  covarianceS(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COVARIANCE.S"), (array1, array2) => webSheetFormulaCovariance(array1, array2, true));
  }

  covarianceP(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COVARIANCE.P"), (array1, array2) => webSheetFormulaCovariance(array1, array2, false));
  }

  covar(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COVAR"), (array1, array2) => webSheetFormulaCovariance(array1, array2, false));
  }

  slope(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SLOPE"), (knownY, knownX) => {
      const regression = webSheetFormulaRegression(knownY, knownX);
      return regression instanceof CellError ? regression : regression.slope;
    });
  }

  avedev(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("AVEDEV"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : webSheetFormulaAverageDeviation(values);
    });
  }

  devsq(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DEVSQ"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : webSheetFormulaDevSq(values);
    });
  }

  geomean(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GEOMEAN"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : webSheetFormulaGeoMean(values);
    });
  }

  harmean(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("HARMEAN"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : webSheetFormulaHarMean(values);
    });
  }

  pearson(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PEARSON"), (array1, array2) => webSheetFormulaCorrelation(array1, array2));
  }

  rsq(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("RSQ"), (knownY, knownX) => {
      const correlation = webSheetFormulaCorrelation(knownY, knownX);
      return correlation instanceof CellError ? correlation : correlation ** 2;
    });
  }

  steyx(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("STEYX"), (knownY, knownX) => webSheetFormulaSteyx(knownY, knownX));
  }

  skew(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SKEW"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : webSheetFormulaSkew(values, false);
    });
  }

  skewP(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SKEW.P"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : webSheetFormulaSkew(values, true);
    });
  }

  normDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NORM.DIST"), (x, mean, standardDev, cumulative) =>
      cumulative ? webSheetFormulaNormalCdf(x, mean, standardDev) : webSheetFormulaNormalPdf(x, mean, standardDev),
    );
  }

  normDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NORMDIST"), (x, mean, standardDev, cumulative) =>
      cumulative ? webSheetFormulaNormalCdf(x, mean, standardDev) : webSheetFormulaNormalPdf(x, mean, standardDev),
    );
  }

  normSDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NORM.S.DIST"), (z, cumulative) =>
      cumulative ? webSheetFormulaNormalCdf(z, 0, 1) : webSheetFormulaNormalPdf(z, 0, 1),
    );
  }

  normSDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NORMSDIST"), (z) => webSheetFormulaNormalCdf(z, 0, 1));
  }

  normInv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NORM.INV"), (probability, mean, standardDev) =>
      webSheetFormulaNormalInv(probability, mean, standardDev),
    );
  }

  normInvLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NORMINV"), (probability, mean, standardDev) =>
      webSheetFormulaNormalInv(probability, mean, standardDev),
    );
  }

  normSInv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NORM.S.INV"), (probability) => webSheetFormulaInverseStandardNormal(probability));
  }

  normSInvLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NORMSINV"), (probability) => webSheetFormulaInverseStandardNormal(probability));
  }

  standardize(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("STANDARDIZE"), (x, mean, standardDev) =>
      standardDev > 0 ? (x - mean) / standardDev : webSheetFormulaError(ErrorType.NUM),
    );
  }

  trimmean(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TRIMMEAN"), (array, percent) =>
      webSheetFormulaTrimMean(array, percent),
    );
  }

  exponDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("EXPON.DIST"), (x, lambda, cumulative) =>
      webSheetFormulaExponDist(x, lambda, cumulative),
    );
  }

  exponDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("EXPONDIST"), (x, lambda, cumulative) =>
      webSheetFormulaExponDist(x, lambda, cumulative),
    );
  }

  poissonDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("POISSON.DIST"), (x, mean, cumulative) =>
      webSheetFormulaPoissonDist(x, mean, cumulative),
    );
  }

  poissonDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("POISSON"), (x, mean, cumulative) =>
      webSheetFormulaPoissonDist(x, mean, cumulative),
    );
  }

  binomDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BINOM.DIST"), (numberS, trials, probabilityS, cumulative) =>
      webSheetFormulaBinomDist(numberS, trials, probabilityS, cumulative),
    );
  }

  binomDistRange(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BINOM.DIST.RANGE"), (trials, probabilityS, numberS, numberS2) =>
      webSheetFormulaBinomDistRange(trials, probabilityS, numberS, numberS2),
    );
  }

  binomDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BINOMDIST"), (numberS, trials, probabilityS, cumulative) =>
      webSheetFormulaBinomDist(numberS, trials, probabilityS, cumulative),
    );
  }

  binomInv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BINOM.INV"), (trials, probabilityS, alpha) =>
      webSheetFormulaBinomInv(trials, probabilityS, alpha),
    );
  }

  critbinom(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CRITBINOM"), (trials, probabilityS, alpha) =>
      webSheetFormulaBinomInv(trials, probabilityS, alpha),
    );
  }

  fisher(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FISHER"), (x) =>
      x > -1 && x < 1 ? 0.5 * Math.log((1 + x) / (1 - x)) : webSheetFormulaError(ErrorType.NUM),
    );
  }

  fisherinv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FISHERINV"), (y) => Math.tanh(y));
  }

  gamma(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GAMMA"), (number) => webSheetFormulaGamma(number));
  }

  gammaln(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GAMMALN"), (number) => {
      const result = webSheetFormulaLogGammaPositive(number);
      return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
    });
  }

  gammalnPrecise(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GAMMALN.PRECISE"), (number) => {
      const result = webSheetFormulaLogGammaPositive(number);
      return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
    });
  }

  gammaDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GAMMA.DIST"), (x, alpha, beta, cumulative) =>
      webSheetFormulaGammaDist(x, alpha, beta, cumulative),
    );
  }

  gammaDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GAMMADIST"), (x, alpha, beta, cumulative) =>
      webSheetFormulaGammaDist(x, alpha, beta, cumulative),
    );
  }

  gammaInv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GAMMA.INV"), (probability, alpha, beta) =>
      webSheetFormulaGammaInv(probability, alpha, beta),
    );
  }

  gammaInvLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GAMMAINV"), (probability, alpha, beta) =>
      webSheetFormulaGammaInv(probability, alpha, beta),
    );
  }

  betaDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BETA.DIST"), (x, alpha, beta, cumulative, lower, upper) =>
      webSheetFormulaBetaDist(x, alpha, beta, cumulative, lower, upper),
    );
  }

  betaDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BETADIST"), (x, alpha, beta, lower, upper) =>
      webSheetFormulaBetaDist(x, alpha, beta, true, lower, upper),
    );
  }

  betaInv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BETA.INV"), (probability, alpha, beta, lower, upper) =>
      webSheetFormulaBetaInv(probability, alpha, beta, lower, upper),
    );
  }

  betaInvLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BETAINV"), (probability, alpha, beta, lower, upper) =>
      webSheetFormulaBetaInv(probability, alpha, beta, lower, upper),
    );
  }

  chisqDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CHISQ.DIST"), (x, degFreedom, cumulative) =>
      webSheetFormulaChiSqDist(x, degFreedom, cumulative),
    );
  }

  chisqInv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CHISQ.INV"), (probability, degFreedom) =>
      webSheetFormulaChiSqInv(probability, degFreedom, false),
    );
  }

  chiDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CHIDIST"), (x, degFreedom) => {
      const cumulative = webSheetFormulaChiSqDist(x, degFreedom, true);
      return cumulative instanceof CellError ? cumulative : 1 - cumulative;
    });
  }

  chiInvLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CHIINV"), (probability, degFreedom) =>
      webSheetFormulaChiSqInv(probability, degFreedom, true),
    );
  }

  weibullDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("WEIBULL.DIST"), (x, alpha, beta, cumulative) =>
      webSheetFormulaWeibullDist(x, alpha, beta, cumulative),
    );
  }

  weibullDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("WEIBULL"), (x, alpha, beta, cumulative) =>
      webSheetFormulaWeibullDist(x, alpha, beta, cumulative),
    );
  }

  hypgeomDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("HYPGEOM.DIST"), (sampleS, numberSample, populationS, numberPop, cumulative) =>
      webSheetFormulaHypergeomDist(sampleS, numberSample, populationS, numberPop, cumulative),
    );
  }

  hypgeomDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("HYPGEOMDIST"), (sampleS, numberSample, populationS, numberPop) =>
      webSheetFormulaHypergeomDist(sampleS, numberSample, populationS, numberPop, false),
    );
  }

  confidenceNorm(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CONFIDENCE.NORM"), (alpha, standardDev, size) =>
      webSheetFormulaConfidenceNorm(alpha, standardDev, size),
    );
  }

  confidenceT(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CONFIDENCE.T"), (alpha, standardDev, size) =>
      webSheetFormulaConfidenceT(alpha, standardDev, size),
    );
  }

  confidenceLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CONFIDENCE"), (alpha, standardDev, size) =>
      webSheetFormulaConfidenceNorm(alpha, standardDev, size),
    );
  }

  fDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("F.DIST"), (x, degFreedom1, degFreedom2, cumulative) =>
      webSheetFormulaFDist(x, degFreedom1, degFreedom2, cumulative),
    );
  }

  fDistRt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("F.DIST.RT"), (x, degFreedom1, degFreedom2) => {
      const cumulative = webSheetFormulaFDist(x, degFreedom1, degFreedom2, true);
      return cumulative instanceof CellError ? cumulative : 1 - cumulative;
    });
  }

  fInv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("F.INV"), (probability, degFreedom1, degFreedom2) =>
      webSheetFormulaFInv(probability, degFreedom1, degFreedom2, false),
    );
  }

  fInvRt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("F.INV.RT"), (probability, degFreedom1, degFreedom2) =>
      webSheetFormulaFInv(probability, degFreedom1, degFreedom2, true),
    );
  }

  fDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FDIST"), (x, degFreedom1, degFreedom2) => {
      const cumulative = webSheetFormulaFDist(x, degFreedom1, degFreedom2, true);
      return cumulative instanceof CellError ? cumulative : 1 - cumulative;
    });
  }

  fInvLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FINV"), (probability, degFreedom1, degFreedom2) =>
      webSheetFormulaFInv(probability, degFreedom1, degFreedom2, true),
    );
  }

  fTest(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("F.TEST"), (array1, array2) =>
      webSheetFormulaFTest(array1, array2),
    );
  }

  fTestLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FTEST"), (array1, array2) =>
      webSheetFormulaFTest(array1, array2),
    );
  }

  zTest(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("Z.TEST"), (array, x, sigma) =>
      webSheetFormulaZTest(array, x, sigma),
    );
  }

  zTestLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ZTEST"), (array, x, sigma) =>
      webSheetFormulaZTest(array, x, sigma),
    );
  }

  chisqTest(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CHISQ.TEST"), (actualRange, expectedRange) =>
      webSheetFormulaChiSqTest(actualRange, expectedRange),
    );
  }

  chitestLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CHITEST"), (actualRange, expectedRange) =>
      webSheetFormulaChiSqTest(actualRange, expectedRange),
    );
  }

  tTest(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("T.TEST"), (array1, array2, tails, type) =>
      webSheetFormulaTTest(array1, array2, tails, type),
    );
  }

  ttestLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TTEST"), (array1, array2, tails, type) =>
      webSheetFormulaTTest(array1, array2, tails, type),
    );
  }

  tDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("T.DIST"), (x, degFreedom, cumulative) =>
      cumulative ? webSheetFormulaStudentTCdf(x, degFreedom) : webSheetFormulaStudentTPdf(x, degFreedom),
    );
  }

  tDistRt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("T.DIST.RT"), (x, degFreedom) => {
      const cumulative = webSheetFormulaStudentTCdf(x, degFreedom);
      return cumulative instanceof CellError ? cumulative : 1 - cumulative;
    });
  }

  tDist2t(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("T.DIST.2T"), (x, degFreedom) =>
      webSheetFormulaStudentTDistTwoTail(x, degFreedom),
    );
  }

  tInv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("T.INV"), (probability, degFreedom) =>
      webSheetFormulaStudentTInv(probability, degFreedom),
    );
  }

  tInv2t(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("T.INV.2T"), (probability, degFreedom) => {
      if (!(probability > 0 && probability < 1)) return webSheetFormulaError(ErrorType.NUM);
      const result = dataAnalysisStudentTInverseTwoTail(probability, degFreedom);
      return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
    });
  }

  tDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TDIST"), (x, degFreedom, tails) => {
      const normalizedTails = Math.trunc(Number(tails));
      if (normalizedTails === 1) {
        const cumulative = webSheetFormulaStudentTCdf(x, degFreedom);
        return cumulative instanceof CellError ? cumulative : 1 - cumulative;
      }
      if (normalizedTails === 2) return webSheetFormulaStudentTDistTwoTail(x, degFreedom);
      return webSheetFormulaError(ErrorType.NUM);
    });
  }

  tDistRtLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TDISTRT"), (x, degFreedom) => {
      const cumulative = webSheetFormulaStudentTCdf(x, degFreedom);
      return cumulative instanceof CellError ? cumulative : 1 - cumulative;
    });
  }

  tDist2tLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TDIST2T"), (x, degFreedom) =>
      webSheetFormulaStudentTDistTwoTail(x, degFreedom),
    );
  }

  tInvLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TINV"), (probability, degFreedom) => {
      const result = dataAnalysisStudentTInverseTwoTail(probability, degFreedom);
      return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
    });
  }

  lognormDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LOGNORM.DIST"), (x, mean, standardDev, cumulative) =>
      webSheetFormulaLogNormDist(x, mean, standardDev, cumulative),
    );
  }

  lognormDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LOGNORMDIST"), (x, mean, standardDev) =>
      webSheetFormulaLogNormDist(x, mean, standardDev, true),
    );
  }

  lognormInv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LOGNORM.INV"), (probability, mean, standardDev) =>
      webSheetFormulaLogNormInv(probability, mean, standardDev),
    );
  }

  lognormInvLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LOGNORMINV"), (probability, mean, standardDev) =>
      webSheetFormulaLogNormInv(probability, mean, standardDev),
    );
  }

  negbinomDist(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NEGBINOM.DIST"), (numberF, numberS, probabilityS, cumulative) =>
      webSheetFormulaNegbinomDist(numberF, numberS, probabilityS, cumulative),
    );
  }

  negbinomDistLegacy(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NEGBINOMDIST"), (numberF, numberS, probabilityS) =>
      webSheetFormulaNegbinomDist(numberF, numberS, probabilityS, false),
    );
  }

  erf(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ERF"), (lowerLimit, upperLimit) => {
      const lower = webSheetFormulaErf(lowerLimit);
      if (upperLimit === WEBSHEET_FORMULA_NA_SENTINEL) return lower;
      return webSheetFormulaErf(upperLimit) - lower;
    });
  }

  erfc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ERFC"), (x) => 1 - webSheetFormulaErf(x));
  }

  erfPrecise(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ERF.PRECISE"), (x) => webSheetFormulaErf(x));
  }

  erfcPrecise(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ERFC.PRECISE"), (x) => 1 - webSheetFormulaErf(x));
  }

  phi(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PHI"), (x) => webSheetFormulaNormalPdf(x, 0, 1));
  }

  gauss(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GAUSS"), (z) => webSheetFormulaNormalCdf(z, 0, 1) - 0.5);
  }

  mod(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MOD"), (number, divisor) => webSheetFormulaMod(number, divisor));
  }

  power(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("POWER"), (number, power) => {
      const result = number ** power;
      return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
    });
  }

  sqrt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SQRT"), (number) =>
      number >= 0 ? Math.sqrt(number) : webSheetFormulaError(ErrorType.NUM),
    );
  }

  sqrtpi(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SQRTPI"), (number) =>
      number >= 0 ? Math.sqrt(number * Math.PI) : webSheetFormulaError(ErrorType.NUM),
    );
  }

  ln(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LN"), (number) =>
      number > 0 ? Math.log(number) : webSheetFormulaError(ErrorType.NUM),
    );
  }

  log(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LOG"), (number, base) => {
      const normalizedBase = base === WEBSHEET_FORMULA_NA_SENTINEL ? 10 : base;
      if (number <= 0 || normalizedBase <= 0 || normalizedBase === 1) return webSheetFormulaError(ErrorType.NUM);
      return Math.log(number) / Math.log(normalizedBase);
    });
  }

  log10(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LOG10"), (number) =>
      number > 0 ? Math.log10(number) : webSheetFormulaError(ErrorType.NUM),
    );
  }

  exp(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("EXP"), (number) => {
      const result = Math.exp(number);
      return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
    });
  }

  fact(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FACT"), (number) => webSheetFormulaFactorialResult(number));
  }

  factdouble(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FACTDOUBLE"), (number) => webSheetFormulaDoubleFactorial(number));
  }

  combin(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COMBIN"), (number, numberChosen) =>
      webSheetFormulaCombinationResult(number, numberChosen),
    );
  }

  combina(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COMBINA"), (number, numberChosen) =>
      webSheetFormulaCombinationA(number, numberChosen),
    );
  }

  permut(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PERMUT"), (number, numberChosen) =>
      webSheetFormulaPermutation(number, numberChosen),
    );
  }

  permutationa(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PERMUTATIONA"), (number, numberChosen) =>
      webSheetFormulaPermutationA(number, numberChosen),
    );
  }

  multinomial(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MULTINOMIAL"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      if (values instanceof CellError) return values;
      if (values.some((value) => value < 0)) return webSheetFormulaError(ErrorType.NUM);
      const ints = values.map((value) => Math.trunc(value));
      const total = ints.reduce((sum, value) => sum + value, 0);
      const numerator = webSheetFormulaFactorialResult(total);
      if (numerator instanceof CellError) return numerator;
      const denominator = ints.reduce((product, value) => product * webSheetFormulaFactorial(value), 1);
      return numerator / denominator;
    });
  }

  gcd(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GCD"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : webSheetFormulaGreatestCommonDivisor(values.map((value) => Math.trunc(value)));
    });
  }

  lcm(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LCM"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : webSheetFormulaLeastCommonMultiple(values.map((value) => Math.trunc(value)));
    });
  }

  pi(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PI"), () => Math.PI);
  }

  sin(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SIN"), (number) => Math.sin(number));
  }

  cos(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COS"), (number) => Math.cos(number));
  }

  tan(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TAN"), (number) => Math.tan(number));
  }

  asin(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ASIN"), (number) =>
      number >= -1 && number <= 1 ? Math.asin(number) : webSheetFormulaError(ErrorType.NUM),
    );
  }

  acos(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ACOS"), (number) =>
      number >= -1 && number <= 1 ? Math.acos(number) : webSheetFormulaError(ErrorType.NUM),
    );
  }

  atan(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ATAN"), (number) => Math.atan(number));
  }

  atan2(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ATAN2"), (xNum, yNum) =>
      xNum === 0 && yNum === 0 ? webSheetFormulaError(ErrorType.DIV_BY_ZERO) : Math.atan2(yNum, xNum),
    );
  }

  degrees(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DEGREES"), (angle) => angle * 180 / Math.PI);
  }

  radians(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("RADIANS"), (angle) => angle * Math.PI / 180);
  }

  sinh(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SINH"), (number) => Math.sinh(number));
  }

  cosh(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COSH"), (number) => Math.cosh(number));
  }

  tanh(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TANH"), (number) => Math.tanh(number));
  }

  asinh(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ASINH"), (number) => Math.asinh(number));
  }

  acosh(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ACOSH"), (number) =>
      number >= 1 ? Math.acosh(number) : webSheetFormulaError(ErrorType.NUM),
    );
  }

  atanh(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ATANH"), (number) =>
      number > -1 && number < 1 ? Math.atanh(number) : webSheetFormulaError(ErrorType.NUM),
    );
  }

  cot(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COT"), (number) => webSheetFormulaTrigReciprocal(number, Math.tan));
  }

  sec(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SEC"), (number) => webSheetFormulaTrigReciprocal(number, Math.cos));
  }

  csc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CSC"), (number) => webSheetFormulaTrigReciprocal(number, Math.sin));
  }

  acot(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ACOT"), (number) => webSheetFormulaAcot(number));
  }

  coth(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COTH"), (number) => webSheetFormulaTrigReciprocal(number, Math.tanh));
  }

  sech(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SECH"), (number) => webSheetFormulaTrigReciprocal(number, Math.cosh));
  }

  csch(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CSCH"), (number) => webSheetFormulaTrigReciprocal(number, Math.sinh));
  }

  seriessum(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SERIESSUM"), (x, n, m, coefficients) =>
      webSheetFormulaSeriesSum(x, n, m, coefficients),
    );
  }

  sumsq(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SUMSQ"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : values.reduce((sum, value) => sum + value * value, 0);
    });
  }

  sumx2my2(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SUMX2MY2"), (arrayX, arrayY) =>
      webSheetFormulaSumPairwiseSquares(arrayX, arrayY, "x2my2"),
    );
  }

  sumx2py2(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SUMX2PY2"), (arrayX, arrayY) =>
      webSheetFormulaSumPairwiseSquares(arrayX, arrayY, "x2py2"),
    );
  }

  sumxmy2(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SUMXMY2"), (arrayX, arrayY) =>
      webSheetFormulaSumPairwiseSquares(arrayX, arrayY, "xmy2"),
    );
  }

  even(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("EVEN"), (number) => webSheetFormulaRoundAwayToParity(number, 0));
  }

  odd(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ODD"), (number) => webSheetFormulaRoundAwayToParity(number, 1));
  }

  isoCeiling(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISO.CEILING"), (number, significance) =>
      webSheetFormulaCeilingPrecise(number, significance),
    );
  }

  maxifs(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MAXIFS"), (maxRange, ...criteriaArgs) => {
      const matched = webSheetFormulaConditionalValues(maxRange, criteriaArgs);
      return matched instanceof CellError ? matched : matched.length ? Math.max(...matched) : 0;
    });
  }

  minifs(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MINIFS"), (minRange, ...criteriaArgs) => {
      const matched = webSheetFormulaConditionalValues(minRange, criteriaArgs);
      return matched instanceof CellError ? matched : matched.length ? Math.min(...matched) : 0;
    });
  }

  modeSngl(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MODE.SNGL"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : webSheetFormulaModeSnglValue(values);
    });
  }

  modeMult(ast, state) {
    const values = webSheetFormulaNumericValues((ast.args || []).map((arg) => this.evaluateAst(arg, state)));
    if (values instanceof CellError) return values;
    const modes = webSheetFormulaModeMultValues(values);
    return modes instanceof CellError ? modes : SimpleRangeValue.onlyValues(modes.map((value) => [value]));
  }

  modeMultArraySize(ast, state) {
    try {
      const values = webSheetFormulaNumericValues((ast.args || []).map((arg) => this.evaluateAst(arg, state)));
      const modes = values instanceof CellError ? [] : webSheetFormulaModeMultValues(values);
      if (Array.isArray(modes) && modes.length) return new ArraySize(1, modes.length);
    } catch {
      // Fall back to input size when formula values are not available during sizing.
    }
    const sizes = (ast.args || []).map((arg) => this.arraySizeForAst(arg, state)).filter(Boolean);
    const height = Math.max(1, sizes.reduce((sum, size) => sum + Math.max(1, size.width) * Math.max(1, size.height), 0));
    return new ArraySize(1, height);
  }

  kurt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("KURT"), (...args) => {
      const values = webSheetFormulaNumericValues(args);
      return values instanceof CellError ? values : webSheetFormulaKurt(values);
    });
  }

  rankEq(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("RANK.EQ"), (number, ref, order) =>
      webSheetFormulaRank(number, ref, order, false),
    );
  }

  rankAvg(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("RANK.AVG"), (number, ref, order) =>
      webSheetFormulaRank(number, ref, order, true),
    );
  }

  percentileInc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PERCENTILE.INC"), (array, k) => {
      const values = webSheetFormulaNumericValues([array]);
      if (values instanceof CellError) return values;
      return webSheetFormulaPercentile(values, k, false);
    });
  }

  percentileExc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PERCENTILE.EXC"), (array, k) => {
      const values = webSheetFormulaNumericValues([array]);
      if (values instanceof CellError) return values;
      return webSheetFormulaPercentile(values, k, true);
    });
  }

  percentRankInc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PERCENTRANK.INC"), (array, x, significance) => {
      const values = webSheetFormulaNumericValues([array]);
      return values instanceof CellError ? values : webSheetFormulaPercentRank(values, x, significance, false);
    });
  }

  percentRankExc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PERCENTRANK.EXC"), (array, x, significance) => {
      const values = webSheetFormulaNumericValues([array]);
      return values instanceof CellError ? values : webSheetFormulaPercentRank(values, x, significance, true);
    });
  }

  quartileInc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("QUARTILE.INC"), (array, quart) => {
      const q = Math.trunc(Number(quart));
      if (q < 0 || q > 4) return webSheetFormulaError(ErrorType.NUM);
      const values = webSheetFormulaNumericValues([array]);
      if (values instanceof CellError) return values;
      return webSheetFormulaPercentile(values, q / 4, false);
    });
  }

  quartileExc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("QUARTILE.EXC"), (array, quart) => {
      const q = Math.trunc(Number(quart));
      if (q < 1 || q > 3) return webSheetFormulaError(ErrorType.NUM);
      const values = webSheetFormulaNumericValues([array]);
      if (values instanceof CellError) return values;
      return webSheetFormulaPercentile(values, q / 4, true);
    });
  }

  prob(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PROB"), (xRange, probRange, lowerLimit, upperLimit) =>
      webSheetFormulaProb(xRange, probRange, lowerLimit, upperLimit),
    );
  }

  frequency(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FREQUENCY"), (dataArray, binsArray) => {
      const counts = webSheetFormulaFrequency(dataArray, binsArray);
      return counts instanceof CellError ? counts : SimpleRangeValue.onlyValues(counts.map((value) => [value]));
    });
  }

  frequencyArraySize(ast, state) {
    const bins = this.arraySizeForAst(ast.args?.[1], state);
    const count = bins ? Math.max(1, bins.width) * Math.max(1, bins.height) + 1 : 1;
    return new ArraySize(1, count);
  }

  linest(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LINEST"), (knownY, knownX, constant, stats) => {
      const fit = webSheetFormulaMultipleRegression(knownY, knownX, { constant });
      if (fit instanceof CellError) return fit;
      return SimpleRangeValue.onlyValues(webSheetFormulaRegressionResultMatrix(fit, { stats }));
    });
  }

  logest(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LOGEST"), (knownY, knownX, constant, stats) => {
      const fit = webSheetFormulaMultipleRegression(knownY, knownX, { constant, logY: true });
      if (fit instanceof CellError) return fit;
      return SimpleRangeValue.onlyValues(webSheetFormulaRegressionResultMatrix(fit, { stats, logOutput: true }));
    });
  }

  linestArraySize(ast, state) {
    const ySize = this.arraySizeForAst(ast.args?.[0], state);
    const xSize = ast.args?.[1] && ast.args[1].type !== "EMPTY" ? this.arraySizeForAst(ast.args[1], state) : null;
    const featureCount = webSheetFormulaRegressionFeatureCountFromSizes(ySize, xSize);
    const stats = webSheetFormulaAstBooleanLiteral(ast.args?.[3], false);
    return new ArraySize(featureCount + 1, stats ? 5 : 1);
  }

  trend(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TREND"), (knownY, knownX, newX, constant) => {
      const fit = webSheetFormulaMultipleRegression(knownY, knownX, { constant });
      if (fit instanceof CellError) return fit;
      const result = webSheetFormulaRegressionPredict(fit, newX);
      return result instanceof CellError ? result : SimpleRangeValue.onlyValues(result);
    });
  }

  growth(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GROWTH"), (knownY, knownX, newX, constant) => {
      const fit = webSheetFormulaMultipleRegression(knownY, knownX, { constant, logY: true });
      if (fit instanceof CellError) return fit;
      const result = webSheetFormulaRegressionPredict(fit, newX, { exp: true });
      return result instanceof CellError ? result : SimpleRangeValue.onlyValues(result);
    });
  }

  trendArraySize(ast, state) {
    const ySize = this.arraySizeForAst(ast.args?.[0], state) || new ArraySize(1, 1);
    const xSize = ast.args?.[1] && ast.args[1].type !== "EMPTY" ? this.arraySizeForAst(ast.args[1], state) : null;
    const newXSize = ast.args?.[2] && ast.args[2].type !== "EMPTY" ? this.arraySizeForAst(ast.args[2], state) : null;
    if (!newXSize) return new ArraySize(Math.max(1, ySize.width), Math.max(1, ySize.height));
    const featureCount = webSheetFormulaRegressionFeatureCountFromSizes(ySize, xSize);
    if (featureCount > 1) {
      if (newXSize.width === featureCount) return new ArraySize(1, Math.max(1, newXSize.height));
      if (newXSize.height === featureCount) return new ArraySize(Math.max(1, newXSize.width), 1);
    }
    return new ArraySize(Math.max(1, newXSize.width), Math.max(1, newXSize.height));
  }

  intercept(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("INTERCEPT"), (knownY, knownX) => {
      const regression = webSheetFormulaRegression(knownY, knownX);
      return regression instanceof CellError ? regression : regression.intercept;
    });
  }

  forecastLinear(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FORECAST.LINEAR"), (x, knownY, knownX) => {
      const regression = webSheetFormulaRegression(knownY, knownX);
      return regression instanceof CellError ? regression : regression.intercept + regression.slope * x;
    });
  }

  dollar(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DOLLAR"), (number, decimals) => {
      const rounded = webSheetFormulaRoundToDecimals(number, decimals);
      const formatted = webSheetFormulaFormatFixed(rounded, decimals, true);
      return rounded < 0 ? `($${formatted.slice(1)})` : `$${formatted}`;
    });
  }

  fixed(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FIXED"), (number, decimals, noCommas) =>
      webSheetFormulaFormatFixed(number, decimals, !noCommas),
    );
  }

  numbervalue(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NUMBERVALUE"), (text, decimalSeparator, groupSeparator) => {
      let source = String(text ?? "").trim();
      const percentCount = (source.match(/%/g) || []).length;
      source = source.replace(/%/g, "");
      const group = String(groupSeparator || ",");
      if (group) source = source.split(group).join("");
      const decimal = String(decimalSeparator || ".");
      if (decimal && decimal !== ".") source = source.replace(decimal, ".");
      const number = Number(source);
      return Number.isFinite(number) ? number / (100 ** percentCount) : webSheetFormulaError(ErrorType.VALUE);
    });
  }

  errorType(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ERROR.TYPE"), (value) => {
      if (!(value instanceof CellError)) return webSheetFormulaError(ErrorType.NA);
      return {
        [ErrorType.DIV_BY_ZERO]: 2,
        [ErrorType.VALUE]: 3,
        [ErrorType.REF]: 4,
        [ErrorType.NAME]: 5,
        [ErrorType.NUM]: 6,
        [ErrorType.NA]: 7,
        [ErrorType.SPILL]: 9,
        [ErrorType.ERROR]: 14,
        [ErrorType.CYCLE]: 14,
      }[value.type] || 14;
    });
  }

  isblank(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISBLANK"), (value) =>
      webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(webSheetFormulaScalar(value))),
    );
  }

  iserror(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISERROR"), (value) => value instanceof CellError);
  }

  iserr(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISERR"), (value) => value instanceof CellError && value.type !== ErrorType.NA);
  }

  isna(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISNA"), (value) => value instanceof CellError && value.type === ErrorType.NA);
  }

  isnumber(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISNUMBER"), (value) =>
      typeof webSheetFormulaRawValue(webSheetFormulaScalar(value)) === "number",
    );
  }

  istext(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISTEXT"), (value) =>
      typeof webSheetFormulaRawValue(webSheetFormulaScalar(value)) === "string",
    );
  }

  isnontext(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISNONTEXT"), (value) =>
      typeof webSheetFormulaRawValue(webSheetFormulaScalar(value)) !== "string",
    );
  }

  islogical(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISLOGICAL"), (value) =>
      typeof webSheetFormulaRawValue(webSheetFormulaScalar(value)) === "boolean",
    );
  }

  iseven(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISEVEN"), (value) => {
      const number = webSheetFormulaNumber(value);
      return number instanceof CellError ? number : Number.isFinite(number) ? Math.trunc(Math.abs(number)) % 2 === 0 : webSheetFormulaError(ErrorType.VALUE);
    });
  }

  isodd(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISODD"), (value) => {
      const number = webSheetFormulaNumber(value);
      return number instanceof CellError ? number : Number.isFinite(number) ? Math.trunc(Math.abs(number)) % 2 === 1 : webSheetFormulaError(ErrorType.VALUE);
    });
  }

  nFunction(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("N"), (value) => webSheetFormulaInfoN(value));
  }

  na() {
    return webSheetFormulaError(ErrorType.NA);
  }

  type(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TYPE"), (value) => {
      if (value instanceof SimpleRangeValue) return 64;
      if (value instanceof CellError) return 16;
      const raw = webSheetFormulaRawValue(value);
      if (typeof raw === "boolean") return 4;
      if (typeof raw === "string") return 2;
      return 1;
    });
  }

  isref(ast, state) {
    if ((ast.args || []).length !== 1) return webSheetFormulaError(ErrorType.NA);
    if (webSheetFormulaReferenceFromAstArg(ast.args[0], state.formulaAddress)) return true;
    return this.runFunction(ast.args, state, this.metadata("ISREF"), (value) => value instanceof SimpleRangeValue);
  }

  isformula(ast, state) {
    if ((ast.args || []).length !== 1) return webSheetFormulaError(ErrorType.NA);
    const ref = webSheetFormulaReferenceCell(webSheetFormulaReferenceFromAstArg(ast.args[0], state.formulaAddress));
    return Boolean(cellHasFormulaContent(ref?.cell));
  }

  formulatext(ast, state) {
    if ((ast.args || []).length !== 1) return webSheetFormulaError(ErrorType.NA);
    const ref = webSheetFormulaReferenceCell(webSheetFormulaReferenceFromAstArg(ast.args[0], state.formulaAddress));
    const formula = ref?.cell?.formula || (typeof ref?.cell?.raw === "string" && ref.cell.raw.startsWith("=") ? ref.cell.raw : "");
    return formula ? String(formula) : webSheetFormulaError(ErrorType.NA);
  }

  sheet(ast, state) {
    const args = ast.args || [];
    if (args.length > 1) return webSheetFormulaError(ErrorType.NA);
    if (!args.length) return webSheetFormulaSheetIndexFromFormulaSheetId(state.formulaAddress?.sheet) + 1;
    const ref = webSheetFormulaReferenceFromAstArg(args[0], state.formulaAddress);
    if (ref) return ref.sheetIndex + 1;
    return this.runFunction(args, state, this.metadata("SHEET"), (value) => {
      const text = String(webSheetFormulaRawValue(value) ?? "").trim();
      const index = state.model?.sheets?.findIndex((sheet) => sheet.name === text) ?? -1;
      return index >= 0 ? index + 1 : webSheetFormulaError(ErrorType.NA);
    });
  }

  sheets(ast, state) {
    const args = ast.args || [];
    if (args.length > 1) return webSheetFormulaError(ErrorType.NA);
    if (!args.length) return state.model?.sheets?.length || 0;
    if (webSheetFormulaReferenceFromAstArg(args[0], state.formulaAddress)) return 1;
    return this.runFunction(args, state, this.metadata("SHEETS"), (value) => {
      if (value instanceof SimpleRangeValue) return 1;
      const text = String(webSheetFormulaRawValue(value) ?? "").trim();
      return (state.model?.sheets || []).some((sheet) => sheet.name === text) ? 1 : webSheetFormulaError(ErrorType.NA);
    });
  }

  cell(ast, state) {
    const args = ast.args || [];
    if (args.length < 1 || args.length > 2) return webSheetFormulaError(ErrorType.NA);
    const infoType = webSheetFormulaText(this.evaluateAst(args[0], state));
    if (infoType instanceof CellError) return infoType;
    const refRange = args[1]
      ? webSheetFormulaReferenceFromAstArg(args[1], state.formulaAddress)
      : {
          sheetIndex: webSheetFormulaSheetIndexFromFormulaSheetId(state.formulaAddress?.sheet),
          top: Number(state.formulaAddress?.row || 0) + 1,
          left: Number(state.formulaAddress?.col || 0) + 1,
          bottom: Number(state.formulaAddress?.row || 0) + 1,
          right: Number(state.formulaAddress?.col || 0) + 1,
        };
    const ref = webSheetFormulaReferenceCell(refRange);
    if (!ref) return webSheetFormulaError(ErrorType.REF);
    const key = String(infoType || "").trim().toLocaleLowerCase();
    const sheetId = getSheetId(ref.sheet.name);
    if (key === "address") return `$${columnName(ref.col)}$${ref.row}`;
    if (key === "col") return ref.col;
    if (key === "row") return ref.row;
    if (key === "contents") return formulaScalarValueForCell(ref.sheet, sheetId, ref.row, ref.col);
    if (key === "filename") return `${webSheetFormulaWorkbookFileName()}${ref.sheet.name || ""}`;
    if (key === "format") return webSheetFormulaCellInfoFormat(ref.cell);
    if (key === "type") return webSheetFormulaCellInfoType(ref.cell);
    if (key === "width") {
      const width = Number(ref.sheet.columns?.[ref.col - 1]?.width || ref.sheet.defaultColumnWidth || columnWidthToPixels(8.43));
      return Math.max(0, Math.round(width / 7));
    }
    if (key === "prefix") return "";
    if (key === "protect") return ref.cell?.protection?.locked === false ? 0 : 1;
    return webSheetFormulaError(ErrorType.VALUE);
  }

  info(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("INFO"), (typeText) => {
      const key = String(typeText || "").trim().toLocaleLowerCase();
      if (key === "system") return "mac";
      if (key === "release") return "WebSheet 0.1.0";
      if (key === "recalc") return "Automatic";
      if (key === "directory") return "";
      if (key === "numfile") return state.model?.sheets?.length || 0;
      if (key === "origin") return "$A:$A";
      if (key === "osversion") {
        if (typeof navigator !== "undefined" && navigator.platform) return navigator.platform;
        return "macOS";
      }
      return webSheetFormulaError(ErrorType.VALUE);
    });
  }

  row(ast, state) {
    const args = ast.args || [];
    if (args.length > 1) return webSheetFormulaError(ErrorType.NA);
    if (!args.length) return Number(state.formulaAddress?.row || 0) + 1;
    const ref = webSheetFormulaReferenceFromAstArg(args[0], state.formulaAddress);
    if (ref) return ref.top;
    return this.runFunction(args, state, this.metadata("ROW"), (value) => {
      if (value instanceof SimpleRangeValue) return 1;
      return webSheetFormulaError(ErrorType.VALUE);
    });
  }

  column(ast, state) {
    const args = ast.args || [];
    if (args.length > 1) return webSheetFormulaError(ErrorType.NA);
    if (!args.length) return Number(state.formulaAddress?.col || 0) + 1;
    const ref = webSheetFormulaReferenceFromAstArg(args[0], state.formulaAddress);
    if (ref) return ref.left;
    return this.runFunction(args, state, this.metadata("COLUMN"), (value) => {
      if (value instanceof SimpleRangeValue) return 1;
      return webSheetFormulaError(ErrorType.VALUE);
    });
  }

  rows(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ROWS"), (array) => {
      const { height } = webSheetFormulaMatrixDimensions(webSheetFormulaMatrix(array));
      return Math.max(1, height);
    });
  }

  columns(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COLUMNS"), (array) => {
      const { width } = webSheetFormulaMatrixDimensions(webSheetFormulaMatrix(array));
      return Math.max(1, width);
    });
  }

  hyperlink(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("HYPERLINK"), (linkLocation, friendlyName) => {
      if (linkLocation instanceof CellError) return linkLocation;
      const fallback = webSheetFormulaText(linkLocation);
      if (friendlyName === WEBSHEET_FORMULA_NA_SENTINEL) return fallback;
      const text = webSheetFormulaText(friendlyName);
      return text instanceof CellError ? text : text;
    });
  }

  tFunction(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("T"), (value) => {
      if (value instanceof CellError) return value;
      const raw = webSheetFormulaRawValue(webSheetFormulaScalar(value));
      return typeof raw === "string" ? raw : "";
    });
  }

  lookup(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("LOOKUP"), (lookupValue, lookupVector, resultVector) => {
      const entries = webSheetFormulaLookupEntries(lookupVector, resultVector);
      return webSheetFormulaApproximateLookup(lookupValue, entries);
    });
  }

  mmult(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MMULT"), (array1, array2) => {
      const result = webSheetFormulaMatrixMultiply(array1, array2);
      return result instanceof CellError ? result : SimpleRangeValue.onlyValues(result);
    });
  }

  mmultArraySize(ast, state) {
    const left = this.arraySizeForAst(ast.args?.[0], state);
    const right = this.arraySizeForAst(ast.args?.[1], state);
    if (!left || !right) return new ArraySize(1, 1);
    return new ArraySize(Math.max(1, right.width), Math.max(1, left.height));
  }

  mdeterm(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MDETERM"), (array) => webSheetFormulaMatrixDeterminant(array));
  }

  minverse(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MINVERSE"), (array) => {
      const result = webSheetFormulaMatrixInverse(array);
      return result instanceof CellError ? result : SimpleRangeValue.onlyValues(result);
    });
  }

  minverseArraySize(ast, state) {
    const size = this.arraySizeForAst(ast.args?.[0], state);
    return size ? new ArraySize(Math.max(1, size.width), Math.max(1, size.height)) : new ArraySize(1, 1);
  }

  munit(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MUNIT"), (dimension) => {
      const result = webSheetFormulaMatrixIdentity(dimension);
      return result instanceof CellError ? result : SimpleRangeValue.onlyValues(result);
    });
  }

  munitArraySize(ast, state) {
    let dimension = webSheetFormulaAstNumberLiteral(ast.args?.[0], 1);
    try {
      const evaluated = ast.args?.[0] ? this.evaluateAst(ast.args[0], state) : dimension;
      const raw = webSheetFormulaRawValue(webSheetFormulaScalar(evaluated));
      if (Number.isFinite(Number(raw))) dimension = Number(raw);
    } catch {
      // Literal fallback is enough for sizing before formula values are available.
    }
    const size = Math.max(1, Math.trunc(Number(dimension) || 1));
    return new ArraySize(size, size);
  }

  xlookup(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("XLOOKUP"), (lookupValue, lookupArray, returnArray, ifNotFound, matchMode, searchMode) =>
      webSheetFormulaXlookupResult(lookupValue, lookupArray, returnArray, ifNotFound, matchMode, searchMode),
    );
  }

  xlookupArraySize(ast, state) {
    const lookup = this.arraySizeForAst(ast.args?.[1], state);
    const result = this.arraySizeForAst(ast.args?.[2], state);
    if (!lookup || !result) return new ArraySize(1, 1);
    if (lookup.height === 1 && lookup.width > 1 && result.height > 1) return new ArraySize(1, Math.max(1, result.height));
    if (lookup.width === 1 && lookup.height > 1 && result.width > 1) return new ArraySize(Math.max(1, result.width), 1);
    if (result.width > 1 && result.height > 1) return new ArraySize(Math.max(1, result.width), 1);
    if (result.width > 1) return new ArraySize(Math.max(1, result.width), 1);
    if (lookup.height === 1 && result.height > 1) return new ArraySize(1, Math.max(1, result.height));
    return new ArraySize(1, 1);
  }

  xmatch(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("XMATCH"), (lookupValue, lookupArray, matchMode, searchMode) => {
      const match = webSheetFormulaLookupMatchIndex(lookupValue, lookupArray, matchMode, searchMode);
      if (match.error) return match.error;
      return match.index >= 0 ? match.index + 1 : webSheetFormulaError(ErrorType.NA);
    });
  }

  filter(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FILTER"), (array, include, ifEmpty) => {
      const fallback = ifEmpty === WEBSHEET_FORMULA_NA_SENTINEL ? undefined : ifEmpty;
      const data = filterFormulaMatrix(webSheetFormulaMatrix(array), webSheetFormulaMatrix(include), fallback);
      return SimpleRangeValue.onlyValues(data);
    });
  }

  sort(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SORT"), (array, sortIndex, sortOrder, byCol) => {
      const data = webSheetFormulaMatrix(array);
      const order = sortOrder < 0 ? -1 : 1;
      const index = Math.max(0, Math.trunc(sortIndex || 1) - 1);
      if (!byCol) {
        const sortedRows = [...data].sort((left, right) => webSheetFormulaCompareForSort(left[index], right[index]) * order);
        return SimpleRangeValue.onlyValues(sortedRows);
      }
      const { width } = webSheetFormulaMatrixDimensions(data);
      const columns = Array.from({ length: width }, (_, colIndex) => data.map((row) => row[colIndex]));
      columns.sort((left, right) => webSheetFormulaCompareForSort(left[index], right[index]) * order);
      const result = data.map((_, rowIndex) => columns.map((column) => column[rowIndex]));
      return SimpleRangeValue.onlyValues(result);
    });
  }

  sortby(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SORTBY"), (array, ...sortArgs) => {
      const descriptors = [];
      for (let index = 0; index < sortArgs.length; index += 2) {
        descriptors.push({
          data: webSheetFormulaMatrix(sortArgs[index]),
          order: Number(sortArgs[index + 1] ?? 1) < 0 ? -1 : 1,
        });
      }
      return SimpleRangeValue.onlyValues(sortByFormulaMatrix(webSheetFormulaMatrix(array), descriptors));
    });
  }

  sortArraySize(ast, state) {
    return webSheetFormulaArraySizeFromFirstArg(ast, state, this);
  }

  unique(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("UNIQUE"), (array, byCol, exactlyOnce) => {
      const data = webSheetFormulaMatrix(array);
      if (!byCol) {
        const counts = new Map();
        data.forEach((row) => counts.set(webSheetFormulaRowKey(row), (counts.get(webSheetFormulaRowKey(row)) || 0) + 1));
        const seen = new Set();
        const result = [];
        for (const row of data) {
          const key = webSheetFormulaRowKey(row);
          if (exactlyOnce && counts.get(key) !== 1) continue;
          if (!exactlyOnce && seen.has(key)) continue;
          seen.add(key);
          result.push(row);
        }
        return SimpleRangeValue.onlyValues(result.length ? result : [[webSheetFormulaError(ErrorType.NA)]]);
      }
      const { width } = webSheetFormulaMatrixDimensions(data);
      const columns = Array.from({ length: width }, (_, colIndex) => data.map((row) => row[colIndex]));
      const counts = new Map();
      columns.forEach((column) => counts.set(webSheetFormulaRowKey(column), (counts.get(webSheetFormulaRowKey(column)) || 0) + 1));
      const seen = new Set();
      const uniqueColumns = [];
      for (const column of columns) {
        const key = webSheetFormulaRowKey(column);
        if (exactlyOnce && counts.get(key) !== 1) continue;
        if (!exactlyOnce && seen.has(key)) continue;
        seen.add(key);
        uniqueColumns.push(column);
      }
      const result = data.map((_, rowIndex) => uniqueColumns.map((column) => column[rowIndex]));
      return SimpleRangeValue.onlyValues(result.length && result[0]?.length ? result : [[webSheetFormulaError(ErrorType.NA)]]);
    });
  }

  uniqueArraySize(ast, state) {
    return webSheetFormulaArraySizeFromFirstArg(ast, state, this);
  }

  sequence(ast, interpreterState) {
    return this.runFunction(ast.args, interpreterState, this.metadata("SEQUENCE"), (rows, columns, start, step) => {
      const height = Math.trunc(Number(rows));
      const width = Math.trunc(Number(columns));
      if (!Number.isFinite(height) || !Number.isFinite(width) || height < 1 || width < 1) {
        return webSheetFormulaError(ErrorType.VALUE);
      }
      const first = Number(start ?? 1);
      const increment = Number(step ?? 1);
      if (!Number.isFinite(first) || !Number.isFinite(increment)) return webSheetFormulaError(ErrorType.VALUE);
      return SimpleRangeValue.onlyValues(sequenceFormulaMatrix(height, width, first, increment));
    });
  }

  sequenceArraySize(ast) {
    const height = Math.max(1, Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[0], 1)));
    const width = Math.max(1, Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[1], 1)));
    return new ArraySize(width, height);
  }

  indirect(ast, interpreterState) {
    return this.runFunction(ast.args, interpreterState, this.metadata("INDIRECT"), (refText) => {
      const baseSheetIndex = interpreterState?.formulaAddress?.sheet ?? 0;
      const baseSheet = state.model?.sheets?.[baseSheetIndex] || state.model?.sheets?.[state.activeSheetIndex];
      const range = decodeFormulaRangeForSheet(refText, baseSheet, baseSheetIndex);
      if (!range) return webSheetFormulaError(ErrorType.REF);
      const data = formulaRangeValueMatrix(range);
      if (range.top === range.bottom && range.left === range.right) return data[0]?.[0] ?? "";
      return SimpleRangeValue.onlyValues(data);
    });
  }

  indirectArraySize(ast, interpreterState) {
    const literal = webSheetFormulaAstStringLiteral(ast.args?.[0]);
    if (!literal) return new ArraySize(1, 1);
    const baseSheetIndex = interpreterState?.formulaAddress?.sheet ?? 0;
    const baseSheet = state.model?.sheets?.[baseSheetIndex] || state.model?.sheets?.[state.activeSheetIndex];
    const range = decodeFormulaRangeForSheet(literal, baseSheet, baseSheetIndex);
    return range ? new ArraySize(range.right - range.left + 1, range.bottom - range.top + 1) : new ArraySize(1, 1);
  }

  take(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TAKE"), (array, rows, columns) =>
      SimpleRangeValue.onlyValues(takeFormulaMatrix(webSheetFormulaMatrix(array), rows, columns)),
    );
  }

  drop(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DROP"), (array, rows, columns) =>
      SimpleRangeValue.onlyValues(dropFormulaMatrix(webSheetFormulaMatrix(array), rows, columns)),
    );
  }

  vstack(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("VSTACK"), (...arrays) =>
      SimpleRangeValue.onlyValues(vstackFormulaMatrices(arrays.map(webSheetFormulaMatrix))),
    );
  }

  hstack(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("HSTACK"), (...arrays) =>
      SimpleRangeValue.onlyValues(hstackFormulaMatrices(arrays.map(webSheetFormulaMatrix))),
    );
  }

  chooserows(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CHOOSEROWS"), (array, ...rows) =>
      SimpleRangeValue.onlyValues(chooseRowsFormulaMatrix(webSheetFormulaMatrix(array), rows)),
    );
  }

  choosecols(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CHOOSECOLS"), (array, ...columns) =>
      SimpleRangeValue.onlyValues(chooseColsFormulaMatrix(webSheetFormulaMatrix(array), columns)),
    );
  }

  tocol(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TOCOL"), (array, ignore, scanByColumn) =>
      SimpleRangeValue.onlyValues(tocolFormulaMatrix(webSheetFormulaMatrix(array), ignore, scanByColumn)),
    );
  }

  torow(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TOROW"), (array, ignore, scanByColumn) =>
      SimpleRangeValue.onlyValues(torowFormulaMatrix(webSheetFormulaMatrix(array), ignore, scanByColumn)),
    );
  }

  wraprows(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("WRAPROWS"), (vector, wrapCount, padWith) =>
      SimpleRangeValue.onlyValues(wrapRowsFormulaMatrix(webSheetFormulaFlattenValues(vector).map(webSheetFormulaRawValue), wrapCount, padWith)),
    );
  }

  wrapcols(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("WRAPCOLS"), (vector, wrapCount, padWith) =>
      SimpleRangeValue.onlyValues(wrapColsFormulaMatrix(webSheetFormulaFlattenValues(vector).map(webSheetFormulaRawValue), wrapCount, padWith)),
    );
  }

  expand(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("EXPAND"), (array, rows, columns, padWith) =>
      SimpleRangeValue.onlyValues(expandFormulaMatrix(webSheetFormulaMatrix(array), rows, columns, padWith)),
    );
  }

  makearray(ast, state) {
    const rows = Math.max(1, Math.trunc(webSheetFormulaLetScalarNumber(webSheetFormulaLetAstValue(ast.args?.[0], state, this, new Map()))));
    const cols = Math.max(1, Math.trunc(webSheetFormulaLetScalarNumber(webSheetFormulaLetAstValue(ast.args?.[1], state, this, new Map()))));
    const lambdaAst = ast.args?.[2];
    const matrix = Array.from({ length: rows }, (_row, rowIndex) =>
      Array.from({ length: cols }, (_col, colIndex) => webSheetFormulaApplyLambda(lambdaAst, [rowIndex + 1, colIndex + 1], state, this)),
    );
    return SimpleRangeValue.onlyValues(matrix);
  }

  byrow(ast, state) {
    const data = webSheetFormulaMatrix(this.evaluateAst(ast.args?.[0], state));
    const lambdaAst = ast.args?.[1];
    return SimpleRangeValue.onlyValues(data.map((row) => [webSheetFormulaApplyLambda(lambdaAst, [SimpleRangeValue.onlyValues([row])], state, this)]));
  }

  bycol(ast, state) {
    const data = webSheetFormulaMatrix(this.evaluateAst(ast.args?.[0], state));
    const width = Math.max(0, ...data.map((row) => row.length));
    const output = Array.from({ length: width }, (_item, colIndex) => {
      const column = data.map((row) => [row[colIndex]]);
      return webSheetFormulaApplyLambda(ast.args?.[1], [SimpleRangeValue.onlyValues(column)], state, this);
    });
    return SimpleRangeValue.onlyValues([output]);
  }

  mapFunction(ast, state) {
    const lambdaAst = ast.args?.[ast.args.length - 1];
    const arrays = (ast.args || []).slice(0, -1).map((arg) => webSheetFormulaMatrix(this.evaluateAst(arg, state)));
    const height = Math.max(0, ...arrays.map((matrix) => matrix.length));
    const width = Math.max(0, ...arrays.flatMap((matrix) => matrix.map((row) => row.length)));
    const output = Array.from({ length: height }, (_row, rowIndex) =>
      Array.from({ length: width }, (_col, colIndex) =>
        webSheetFormulaApplyLambda(lambdaAst, arrays.map((matrix) => matrix[rowIndex]?.[colIndex]), state, this),
      ),
    );
    return SimpleRangeValue.onlyValues(output);
  }

  reduce(ast, state) {
    let accumulator = webSheetFormulaLetAstValue(ast.args?.[0], state, this, new Map());
    const values = webSheetFormulaFlattenValues(this.evaluateAst(ast.args?.[1], state));
    const lambdaAst = ast.args?.[2];
    for (const value of values) {
      accumulator = webSheetFormulaApplyLambda(lambdaAst, [accumulator, value], state, this);
      if (accumulator instanceof CellError) return accumulator;
    }
    return accumulator;
  }

  scan(ast, state) {
    let accumulator = webSheetFormulaLetAstValue(ast.args?.[0], state, this, new Map());
    const data = webSheetFormulaMatrix(this.runFunction([ast.args?.[1]], state, { parameters: [WEBSHEET_FORMULA_RANGE] }, (array) => array));
    const lambdaAst = ast.args?.[2];
    const output = data.map((row) => row.map((value) => {
      accumulator = webSheetFormulaApplyLambda(lambdaAst, [accumulator, value], state, this);
      return accumulator;
    }));
    return SimpleRangeValue.onlyValues(output);
  }

  dynamicArraySize(ast, state) {
    return webSheetFormulaArraySizeFromFirstArg(ast, state, this);
  }

  scanArraySize(ast, state) {
    const size = this.arraySizeForAst(ast.args?.[1], state);
    return size?.isRef ? new ArraySize(Math.max(1, size.width), Math.max(1, size.height)) : size;
  }

  takeArraySize(ast, state) {
    const base = this.arraySizeForAst(ast.args?.[0], state);
    const rows = Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[1], base.height));
    const cols = Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[2], base.width));
    return new ArraySize(Math.max(1, Math.min(base.width, Math.abs(cols || base.width))), Math.max(1, Math.min(base.height, Math.abs(rows || base.height))));
  }

  dropArraySize(ast, state) {
    const base = this.arraySizeForAst(ast.args?.[0], state);
    const rows = Math.abs(Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[1], 0)));
    const cols = Math.abs(Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[2], 0)));
    return new ArraySize(Math.max(1, base.width - cols), Math.max(1, base.height - rows));
  }

  vstackArraySize(ast, state) {
    const sizes = ast.args.map((arg) => this.arraySizeForAst(arg, state));
    return new ArraySize(Math.max(1, ...sizes.map((size) => size.width)), Math.max(1, sizes.reduce((sum, size) => sum + size.height, 0)));
  }

  hstackArraySize(ast, state) {
    const sizes = ast.args.map((arg) => this.arraySizeForAst(arg, state));
    return new ArraySize(Math.max(1, sizes.reduce((sum, size) => sum + size.width, 0)), Math.max(1, ...sizes.map((size) => size.height)));
  }

  chooserowsArraySize(ast, state) {
    const base = this.arraySizeForAst(ast.args?.[0], state);
    return new ArraySize(Math.max(1, base.width), Math.max(1, ast.args.length - 1));
  }

  choosecolsArraySize(ast, state) {
    const base = this.arraySizeForAst(ast.args?.[0], state);
    return new ArraySize(Math.max(1, ast.args.length - 1), Math.max(1, base.height));
  }

  tocolArraySize(ast, state) {
    const base = this.arraySizeForAst(ast.args?.[0], state);
    return new ArraySize(1, Math.max(1, base.width * base.height));
  }

  torowArraySize(ast, state) {
    const base = this.arraySizeForAst(ast.args?.[0], state);
    return new ArraySize(Math.max(1, base.width * base.height), 1);
  }

  wraprowsArraySize(ast, state) {
    const base = this.arraySizeForAst(ast.args?.[0], state);
    const width = Math.max(1, Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[1], 1)));
    return new ArraySize(width, Math.max(1, Math.ceil((base.width * base.height) / width)));
  }

  wrapcolsArraySize(ast, state) {
    const base = this.arraySizeForAst(ast.args?.[0], state);
    const height = Math.max(1, Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[1], 1)));
    return new ArraySize(Math.max(1, Math.ceil((base.width * base.height) / height)), height);
  }

  expandArraySize(ast, state) {
    const base = this.arraySizeForAst(ast.args?.[0], state);
    const height = Math.max(base.height, Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[1], base.height)));
    const width = Math.max(base.width, Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[2], base.width)));
    return new ArraySize(width, height);
  }

  makearrayArraySize(ast) {
    return webSheetFormulaMakeArraySize(ast);
  }

  byrowArraySize(ast, state) {
    return webSheetFormulaByRowArraySize(ast, state, this);
  }

  bycolArraySize(ast, state) {
    return webSheetFormulaByColArraySize(ast, state, this);
  }

  convert(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CONVERT"), (number, fromUnit, toUnit) => {
      const fromKey = webSheetFormulaConvertUnitKey(fromUnit);
      const toKey = webSheetFormulaConvertUnitKey(toUnit);
      const temperature = webSheetFormulaConvertTemperature(number, fromKey, toKey);
      if (temperature != null) return temperature;
      const from = WEBSHEET_CONVERT_UNITS.get(fromKey);
      const to = WEBSHEET_CONVERT_UNITS.get(toKey);
      if (!from || !to || from[0] !== to[0]) return webSheetFormulaError(ErrorType.NA);
      return (number * from[1]) / to[1];
    });
  }

  networkdays(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NETWORKDAYS"), (startDate, endDate, holidays) =>
      webSheetFormulaNetworkDays(startDate, endDate, undefined, holidays),
    );
  }

  networkdaysIntl(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NETWORKDAYS.INTL"), (startDate, endDate, weekend, holidays) =>
      webSheetFormulaNetworkDays(startDate, endDate, weekend, holidays),
    );
  }

  workday(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("WORKDAY"), (startDate, days, holidays) =>
      webSheetFormulaWorkday(startDate, days, undefined, holidays),
    );
  }

  workdayIntl(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("WORKDAY.INTL"), (startDate, days, weekend, holidays) =>
      webSheetFormulaWorkday(startDate, days, weekend, holidays),
    );
  }

  datevalue(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DATEVALUE"), (dateText) => webSheetFormulaDateValue(dateText));
  }

  timevalue(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TIMEVALUE"), (timeText) => webSheetFormulaTimeValue(timeText));
  }

  hour(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("HOUR"), (serialNumber) => webSheetFormulaTimeUnit(serialNumber, "hour"));
  }

  minute(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MINUTE"), (serialNumber) => webSheetFormulaTimeUnit(serialNumber, "minute"));
  }

  second(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SECOND"), (serialNumber) => webSheetFormulaTimeUnit(serialNumber, "second"));
  }

  weekday(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("WEEKDAY"), (serialNumber, returnType) =>
      webSheetFormulaWeekdayValue(serialNumber, returnType),
    );
  }

  weeknum(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("WEEKNUM"), (serialNumber, returnType) =>
      webSheetFormulaWeeknumValue(serialNumber, returnType),
    );
  }

  isoweeknum(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISOWEEKNUM"), (serialNumber) =>
      webSheetFormulaIsoWeeknumValue(serialNumber),
    );
  }

  edate(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("EDATE"), (startDate, months) =>
      webSheetFormulaAddMonths(startDate, months, false),
    );
  }

  eomonth(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("EOMONTH"), (startDate, months) =>
      webSheetFormulaAddMonths(startDate, months, true),
    );
  }

  days(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DAYS"), (endDate, startDate) => {
      const start = webSheetFormulaDateSerial(startDate);
      const end = webSheetFormulaDateSerial(endDate);
      if (start instanceof CellError) return start;
      if (end instanceof CellError) return end;
      return end - start;
    });
  }

  days360(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DAYS360"), (startDate, endDate, method) =>
      webSheetFormulaDays360(startDate, endDate, method),
    );
  }

  yearfrac(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("YEARFRAC"), (startDate, endDate, basis) => {
      const mode = Math.trunc(Number(basis) || 0);
      return mode < 0 || mode > 4 ? webSheetFormulaError(ErrorType.NUM) : webSheetFormulaYearFrac(startDate, endDate, mode);
    });
  }

  datedif(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DATEDIF"), (startDate, endDate, unit) => {
      const start = webSheetFormulaDateParts(startDate);
      const end = webSheetFormulaDateParts(endDate);
      if (start instanceof CellError) return start;
      if (end instanceof CellError) return end;
      const value = excelDatedif(start, end, String(unit || "").toUpperCase());
      return value == null ? webSheetFormulaError(ErrorType.NUM) : value;
    });
  }

  ceilingMath(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CEILING.MATH"), (number, significance, mode) =>
      webSheetFormulaCeilingMath(number, significance, mode),
    );
  }

  floorMath(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FLOOR.MATH"), (number, significance, mode) =>
      webSheetFormulaFloorMath(number, significance, mode),
    );
  }

  mround(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MROUND"), (number, multiple) =>
      webSheetFormulaMround(number, multiple),
    );
  }

  ceilingPrecise(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CEILING.PRECISE"), (number, significance) =>
      webSheetFormulaCeilingPrecise(number, significance),
    );
  }

  floorPrecise(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FLOOR.PRECISE"), (number, significance) =>
      webSheetFormulaFloorPrecise(number, significance),
    );
  }

  quotient(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("QUOTIENT"), (numerator, denominator) => {
      const left = Number(numerator);
      const right = Number(denominator);
      if (!Number.isFinite(left) || !Number.isFinite(right)) return webSheetFormulaError(ErrorType.VALUE);
      return right === 0 ? webSheetFormulaError(ErrorType.DIV_BY_ZERO) : Math.trunc(left / right);
    });
  }

  sign(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SIGN"), (number) => Math.sign(Number(number)) || 0);
  }

  pmt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PMT"), (rate, nper, pv, fv, type) =>
      webSheetFormulaPayment(rate, nper, pv, fv, type),
    );
  }

  pv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PV"), (rate, nper, pmt, fv, type) =>
      webSheetFormulaPresentValue(rate, nper, pmt, fv, type),
    );
  }

  fv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FV"), (rate, nper, pmt, pv, type) =>
      webSheetFormulaFutureValue(rate, nper, pmt, pv, type),
    );
  }

  nper(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NPER"), (rate, pmt, pv, fv, type) =>
      webSheetFormulaNper(rate, pmt, pv, fv, type),
    );
  }

  rate(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("RATE"), (nper, pmt, pv, fv, type, guess) =>
      webSheetFormulaRate(nper, pmt, pv, fv, type, guess),
    );
  }

  npv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NPV"), (rate, ...values) => webSheetFormulaNpv(rate, values));
  }

  fvSchedule(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("FVSCHEDULE"), (principal, schedule) =>
      webSheetFormulaFvSchedule(principal, schedule),
    );
  }

  xnpv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("XNPV"), (rate, values, dates) =>
      webSheetFormulaXnpv(rate, values, dates),
    );
  }

  xirr(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("XIRR"), (values, dates, guess) =>
      webSheetFormulaXirr(values, dates, guess),
    );
  }

  irr(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IRR"), (values, guess) => webSheetFormulaIrr(values, guess));
  }

  mirr(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MIRR"), (values, financeRate, reinvestRate) =>
      webSheetFormulaMirr(values, financeRate, reinvestRate),
    );
  }

  db(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DB"), (cost, salvage, life, period, month) =>
      webSheetFormulaDb(cost, salvage, life, period, month),
    );
  }

  ddb(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DDB"), (cost, salvage, life, period, factor) =>
      webSheetFormulaDdb(cost, salvage, life, period, factor),
    );
  }

  vdb(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("VDB"), (cost, salvage, life, startPeriod, endPeriod, factor, noSwitch) =>
      webSheetFormulaVdb(cost, salvage, life, startPeriod, endPeriod, factor, noSwitch),
    );
  }

  amorlinc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("AMORLINC"), (cost, datePurchased, firstPeriod, salvage, period, rate, basis) =>
      webSheetFormulaAmorLinc(cost, datePurchased, firstPeriod, salvage, period, rate, basis),
    );
  }

  amordegrc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("AMORDEGRC"), (cost, datePurchased, firstPeriod, salvage, period, rate, basis) =>
      webSheetFormulaAmorDegrc(cost, datePurchased, firstPeriod, salvage, period, rate, basis),
    );
  }

  sln(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SLN"), (cost, salvage, life) =>
      webSheetFormulaSln(cost, salvage, life),
    );
  }

  syd(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("SYD"), (cost, salvage, life, per) =>
      webSheetFormulaSyd(cost, salvage, life, per),
    );
  }

  dollarde(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DOLLARDE"), (fractionalDollar, fraction) =>
      webSheetFormulaDollarDe(fractionalDollar, fraction),
    );
  }

  dollarfr(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DOLLARFR"), (decimalDollar, fraction) =>
      webSheetFormulaDollarFr(decimalDollar, fraction),
    );
  }

  rri(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("RRI"), (nper, pv, fv) =>
      webSheetFormulaRri(nper, pv, fv),
    );
  }

  pduration(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PDURATION"), (rate, pv, fv) =>
      webSheetFormulaPduration(rate, pv, fv),
    );
  }

  ispmt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ISPMT"), (rate, per, nper, pv) =>
      webSheetFormulaIspmt(rate, per, nper, pv),
    );
  }

  tbillEq(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TBILLEQ"), (settlement, maturity, discount) =>
      webSheetFormulaTbillEq(settlement, maturity, discount),
    );
  }

  tbillPrice(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TBILLPRICE"), (settlement, maturity, discount) =>
      webSheetFormulaTbillPrice(settlement, maturity, discount),
    );
  }

  tbillYield(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TBILLYIELD"), (settlement, maturity, price) =>
      webSheetFormulaTbillYield(settlement, maturity, price),
    );
  }

  disc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DISC"), (settlement, maturity, pr, redemption, basis) =>
      webSheetFormulaDisc(settlement, maturity, pr, redemption, basis),
    );
  }

  priceDisc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PRICEDISC"), (settlement, maturity, discount, redemption, basis) =>
      webSheetFormulaPriceDisc(settlement, maturity, discount, redemption, basis),
    );
  }

  yieldDisc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("YIELDDISC"), (settlement, maturity, pr, redemption, basis) =>
      webSheetFormulaYieldDisc(settlement, maturity, pr, redemption, basis),
    );
  }

  intrate(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("INTRATE"), (settlement, maturity, investment, redemption, basis) =>
      webSheetFormulaIntrate(settlement, maturity, investment, redemption, basis),
    );
  }

  received(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("RECEIVED"), (settlement, maturity, investment, discount, basis) =>
      webSheetFormulaReceived(settlement, maturity, investment, discount, basis),
    );
  }

  accrint(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ACCRINT"), (issue, firstInterest, settlement, rate, par, frequency, basis, calcMethod) =>
      webSheetFormulaAccrint(issue, firstInterest, settlement, rate, par, frequency, basis, calcMethod),
    );
  }

  accrintM(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ACCRINTM"), (issue, settlement, rate, par, basis) =>
      webSheetFormulaAccrintM(issue, settlement, rate, par, basis),
    );
  }

  priceMat(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PRICEMAT"), (settlement, maturity, issue, rate, yld, basis) =>
      webSheetFormulaPriceMat(settlement, maturity, issue, rate, yld, basis),
    );
  }

  yieldMat(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("YIELDMAT"), (settlement, maturity, issue, rate, pr, basis) =>
      webSheetFormulaYieldMat(settlement, maturity, issue, rate, pr, basis),
    );
  }

  coupDayBs(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COUPDAYBS"), (settlement, maturity, frequency, basis) =>
      webSheetFormulaCoupDayBs(settlement, maturity, frequency, basis),
    );
  }

  coupDays(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COUPDAYS"), (settlement, maturity, frequency, basis) =>
      webSheetFormulaCoupDays(settlement, maturity, frequency, basis),
    );
  }

  coupDaysNc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COUPDAYSNC"), (settlement, maturity, frequency, basis) =>
      webSheetFormulaCoupDaysNc(settlement, maturity, frequency, basis),
    );
  }

  coupNcd(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COUPNCD"), (settlement, maturity, frequency, basis) =>
      webSheetFormulaCoupNcd(settlement, maturity, frequency, basis),
    );
  }

  coupPcd(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COUPPCD"), (settlement, maturity, frequency, basis) =>
      webSheetFormulaCoupPcd(settlement, maturity, frequency, basis),
    );
  }

  coupNum(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COUPNUM"), (settlement, maturity, frequency, basis) =>
      webSheetFormulaCoupNum(settlement, maturity, frequency, basis),
    );
  }

  price(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PRICE"), (settlement, maturity, rate, yld, redemption, frequency, basis) =>
      webSheetFormulaSecurityPrice(settlement, maturity, rate, yld, redemption, frequency, basis),
    );
  }

  yieldSecurity(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("YIELD"), (settlement, maturity, rate, pr, redemption, frequency, basis) =>
      webSheetFormulaSecurityYield(settlement, maturity, rate, pr, redemption, frequency, basis),
    );
  }

  duration(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DURATION"), (settlement, maturity, coupon, yld, frequency, basis) =>
      webSheetFormulaSecurityDuration(settlement, maturity, coupon, yld, frequency, basis, false),
    );
  }

  mduration(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MDURATION"), (settlement, maturity, coupon, yld, frequency, basis) =>
      webSheetFormulaSecurityDuration(settlement, maturity, coupon, yld, frequency, basis, true),
    );
  }

  ipmt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IPMT"), (rate, per, nper, pv, fv, type) =>
      webSheetFormulaIpmt(rate, per, nper, pv, fv, type),
    );
  }

  ppmt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("PPMT"), (rate, per, nper, pv, fv, type) =>
      webSheetFormulaPpmt(rate, per, nper, pv, fv, type),
    );
  }

  cumipmt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CUMIPMT"), (rate, nper, pv, startPeriod, endPeriod, type) => {
      const start = Math.trunc(Number(startPeriod));
      const end = Math.trunc(Number(endPeriod));
      if (start < 1 || end < start) return webSheetFormulaError(ErrorType.NUM);
      let total = 0;
      for (let period = start; period <= end; period += 1) {
        const value = webSheetFormulaIpmt(rate, period, nper, pv, 0, type);
        if (value instanceof CellError) return value;
        total += value;
      }
      return total;
    });
  }

  cumprinc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CUMPRINC"), (rate, nper, pv, startPeriod, endPeriod, type) => {
      const start = Math.trunc(Number(startPeriod));
      const end = Math.trunc(Number(endPeriod));
      if (start < 1 || end < start) return webSheetFormulaError(ErrorType.NUM);
      let total = 0;
      for (let period = start; period <= end; period += 1) {
        const value = webSheetFormulaPpmt(rate, period, nper, pv, 0, type);
        if (value instanceof CellError) return value;
        total += value;
      }
      return total;
    });
  }

  effect(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("EFFECT"), (nominalRate, npery) => {
      const rate = Number(nominalRate);
      const periods = Math.trunc(Number(npery));
      return Number.isFinite(rate) && periods >= 1 ? ((1 + rate / periods) ** periods) - 1 : webSheetFormulaError(ErrorType.NUM);
    });
  }

  nominal(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("NOMINAL"), (effectRate, npery) => {
      const rate = Number(effectRate);
      const periods = Math.trunc(Number(npery));
      return Number.isFinite(rate) && rate > -1 && periods >= 1 ? periods * (((1 + rate) ** (1 / periods)) - 1) : webSheetFormulaError(ErrorType.NUM);
    });
  }

  roman(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ROMAN"), (number, form) => {
      const mode = Math.trunc(Number(form || 0));
      return mode < 0 || mode > 4 ? webSheetFormulaError(ErrorType.VALUE) : webSheetFormulaRoman(number);
    });
  }

  arabic(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("ARABIC"), (text) => webSheetFormulaArabic(text));
  }

  baseFunction(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BASE"), (number, radix, minLength) =>
      webSheetFormulaBase(number, radix, minLength),
    );
  }

  decimal(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DECIMAL"), (text, radix) => webSheetFormulaDecimal(text, radix));
  }

  dec2bin(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DEC2BIN"), (number, places) => webSheetFormulaDecimalToSignedBase(number, 2, places));
  }

  dec2oct(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DEC2OCT"), (number, places) => webSheetFormulaDecimalToSignedBase(number, 8, places));
  }

  dec2hex(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DEC2HEX"), (number, places) => webSheetFormulaDecimalToSignedBase(number, 16, places));
  }

  bin2dec(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BIN2DEC"), (number) => webSheetFormulaConvertBase(number, 2, 10));
  }

  bin2oct(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BIN2OCT"), (number, places) => webSheetFormulaConvertBase(number, 2, 8, places));
  }

  bin2hex(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BIN2HEX"), (number, places) => webSheetFormulaConvertBase(number, 2, 16, places));
  }

  oct2bin(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("OCT2BIN"), (number, places) => webSheetFormulaConvertBase(number, 8, 2, places));
  }

  oct2dec(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("OCT2DEC"), (number) => webSheetFormulaConvertBase(number, 8, 10));
  }

  oct2hex(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("OCT2HEX"), (number, places) => webSheetFormulaConvertBase(number, 8, 16, places));
  }

  hex2bin(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("HEX2BIN"), (number, places) => webSheetFormulaConvertBase(number, 16, 2, places));
  }

  hex2oct(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("HEX2OCT"), (number, places) => webSheetFormulaConvertBase(number, 16, 8, places));
  }

  hex2dec(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("HEX2DEC"), (number) => webSheetFormulaConvertBase(number, 16, 10));
  }

  bitand(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BITAND"), (number1, number2) =>
      webSheetFormulaBitwise(number1, number2, (left, right) => left & right),
    );
  }

  bitor(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BITOR"), (number1, number2) =>
      webSheetFormulaBitwise(number1, number2, (left, right) => left | right),
    );
  }

  bitxor(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BITXOR"), (number1, number2) =>
      webSheetFormulaBitwise(number1, number2, (left, right) => left ^ right),
    );
  }

  bitlshift(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BITLSHIFT"), (number, shiftAmount) =>
      webSheetFormulaBitShift(number, shiftAmount, 1),
    );
  }

  bitrshift(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BITRSHIFT"), (number, shiftAmount) =>
      webSheetFormulaBitShift(number, shiftAmount, -1),
    );
  }

  delta(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DELTA"), (number1, number2) => (number1 === number2 ? 1 : 0));
  }

  gestep(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("GESTEP"), (number, step) => (number >= step ? 1 : 0));
  }

  unichar(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("UNICHAR"), (number) => {
      const codePoint = Math.trunc(Number(number));
      try {
        return codePoint > 0 ? String.fromCodePoint(codePoint) : webSheetFormulaError(ErrorType.VALUE);
      } catch {
        return webSheetFormulaError(ErrorType.VALUE);
      }
    });
  }

  unicode(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("UNICODE"), (text) => {
      const value = String(text ?? "");
      return value ? value.codePointAt(0) : webSheetFormulaError(ErrorType.VALUE);
    });
  }

  complexFunction(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("COMPLEX"), (realNum, iNum, suffix) =>
      webSheetFormulaComplexText(realNum, iNum, suffix),
    );
  }

  imabs(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMABS"), (value) => webSheetFormulaImabs(value));
  }

  imreal(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMREAL"), (value) => {
      const complex = webSheetFormulaComplex(value);
      return complex instanceof CellError ? complex : complex.real;
    });
  }

  imaginary(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMAGINARY"), (value) => {
      const complex = webSheetFormulaComplex(value);
      return complex instanceof CellError ? complex : complex.imag;
    });
  }

  imsum(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMSUM"), (...values) => webSheetFormulaImsum(values));
  }

  improduct(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMPRODUCT"), (...values) => webSheetFormulaImproduct(values));
  }

  imsub(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMSUB"), (left, right) => {
      const a = webSheetFormulaComplex(left);
      const b = webSheetFormulaComplex(right);
      if (a instanceof CellError) return a;
      if (b instanceof CellError) return b;
      return webSheetFormulaComplexTextOrError({ real: a.real - b.real, imag: a.imag - b.imag });
    });
  }

  imdiv(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMDIV"), (left, right) => {
      const a = webSheetFormulaComplex(left);
      const b = webSheetFormulaComplex(right);
      if (a instanceof CellError) return a;
      if (b instanceof CellError) return b;
      const denominator = b.real ** 2 + b.imag ** 2;
      if (!denominator) return webSheetFormulaError(ErrorType.NUM);
      return dataAnalysisComplexText((a.real * b.real + a.imag * b.imag) / denominator, (a.imag * b.real - a.real * b.imag) / denominator);
    });
  }

  imconjugate(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMCONJUGATE"), (value) => {
      const complex = webSheetFormulaComplex(value);
      return complex instanceof CellError ? complex : dataAnalysisComplexText(complex.real, -complex.imag);
    });
  }

  imargument(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMARGUMENT"), (value) => {
      const complex = webSheetFormulaComplex(value);
      if (complex instanceof CellError) return complex;
      return complex.real === 0 && complex.imag === 0 ? webSheetFormulaError(ErrorType.DIV_BY_ZERO) : Math.atan2(complex.imag, complex.real);
    });
  }

  imexp(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMEXP"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => {
        const magnitude = Math.exp(real);
        return { real: magnitude * Math.cos(imag), imag: magnitude * Math.sin(imag) };
      }),
    );
  }

  imln(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMLN"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => {
        const magnitude = Math.hypot(real, imag);
        return magnitude ? { real: Math.log(magnitude), imag: Math.atan2(imag, real) } : webSheetFormulaError(ErrorType.NUM);
      }),
    );
  }

  imlog10(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMLOG10"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => {
        const magnitude = Math.hypot(real, imag);
        return magnitude ? { real: Math.log10(magnitude), imag: Math.atan2(imag, real) / Math.LN10 } : webSheetFormulaError(ErrorType.NUM);
      }),
    );
  }

  imlog2(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMLOG2"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => {
        const magnitude = Math.hypot(real, imag);
        return magnitude ? { real: Math.log2(magnitude), imag: Math.atan2(imag, real) / Math.LN2 } : webSheetFormulaError(ErrorType.NUM);
      }),
    );
  }

  impowerFunction(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMPOWER"), (value, number) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => {
        const exponent = Number(number);
        const magnitude = Math.hypot(real, imag);
        if (!Number.isFinite(exponent) || magnitude === 0) return webSheetFormulaError(ErrorType.NUM);
        const powered = magnitude ** exponent;
        const angle = Math.atan2(imag, real) * exponent;
        return { real: powered * Math.cos(angle), imag: powered * Math.sin(angle) };
      }),
    );
  }

  imsqrt(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMSQRT"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => {
        const magnitude = Math.hypot(real, imag);
        const angle = Math.atan2(imag, real) / 2;
        return { real: Math.sqrt(magnitude) * Math.cos(angle), imag: Math.sqrt(magnitude) * Math.sin(angle) };
      }),
    );
  }

  imsin(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMSIN"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => ({ real: Math.sin(real) * Math.cosh(imag), imag: Math.cos(real) * Math.sinh(imag) })),
    );
  }

  imcos(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMCOS"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => ({ real: Math.cos(real) * Math.cosh(imag), imag: -Math.sin(real) * Math.sinh(imag) })),
    );
  }

  imtan(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMTAN"), (value) => {
      const sin = webSheetFormulaComplex(value);
      if (sin instanceof CellError) return sin;
      const numerator = { real: Math.sin(sin.real) * Math.cosh(sin.imag), imag: Math.cos(sin.real) * Math.sinh(sin.imag) };
      const denominator = { real: Math.cos(sin.real) * Math.cosh(sin.imag), imag: -Math.sin(sin.real) * Math.sinh(sin.imag) };
      const result = webSheetFormulaComplexDivideValues(numerator, denominator);
      return webSheetFormulaComplexTextOrError(result);
    });
  }

  imsec(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMSEC"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) =>
        webSheetFormulaComplexDivideValues({ real: 1, imag: 0 }, { real: Math.cos(real) * Math.cosh(imag), imag: -Math.sin(real) * Math.sinh(imag) })),
    );
  }

  imcsc(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMCSC"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) =>
        webSheetFormulaComplexDivideValues({ real: 1, imag: 0 }, { real: Math.sin(real) * Math.cosh(imag), imag: Math.cos(real) * Math.sinh(imag) })),
    );
  }

  imcot(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMCOT"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => {
        const sin = { real: Math.sin(real) * Math.cosh(imag), imag: Math.cos(real) * Math.sinh(imag) };
        const cos = { real: Math.cos(real) * Math.cosh(imag), imag: -Math.sin(real) * Math.sinh(imag) };
        return webSheetFormulaComplexDivideValues(cos, sin);
      }),
    );
  }

  imsinh(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMSINH"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => ({ real: Math.sinh(real) * Math.cos(imag), imag: Math.cosh(real) * Math.sin(imag) })),
    );
  }

  imcosh(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMCOSH"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) => ({ real: Math.cosh(real) * Math.cos(imag), imag: Math.sinh(real) * Math.sin(imag) })),
    );
  }

  imsech(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMSECH"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) =>
        webSheetFormulaComplexDivideValues({ real: 1, imag: 0 }, { real: Math.cosh(real) * Math.cos(imag), imag: Math.sinh(real) * Math.sin(imag) })),
    );
  }

  imcsch(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IMCSCH"), (value) =>
      webSheetFormulaComplexUnary(value, ({ real, imag }) =>
        webSheetFormulaComplexDivideValues({ real: 1, imag: 0 }, { real: Math.sinh(real) * Math.cos(imag), imag: Math.cosh(real) * Math.sin(imag) })),
    );
  }

  besseli(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BESSELI"), (x, n) => webSheetFormulaBesselI(x, n));
  }

  besselj(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BESSELJ"), (x, n) => webSheetFormulaBesselJ(x, n));
  }

  besselk(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BESSELK"), (x, n) => webSheetFormulaBesselK(x, n));
  }

  bessely(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("BESSELY"), (x, n) => webSheetFormulaBesselY(x, n));
  }

  dsum(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DSUM"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "sum"),
    );
  }

  dcount(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DCOUNT"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "count"),
    );
  }

  dcounta(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DCOUNTA"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "counta"),
    );
  }

  daverage(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DAVERAGE"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "average"),
    );
  }

  dmax(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DMAX"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "max"),
    );
  }

  dmin(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DMIN"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "min"),
    );
  }

  dproduct(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DPRODUCT"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "product"),
    );
  }

  dstdev(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DSTDEV"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "stdev"),
    );
  }

  dstdevp(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DSTDEVP"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "stdevp"),
    );
  }

  dvar(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DVAR"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "var"),
    );
  }

  dvarp(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("DVARP"), (database, field, criteria) =>
      webSheetFormulaDatabaseAggregate(database, field, criteria, "varp"),
    );
  }
};
}

export {
  createWebSheetFormulaPluginClass,
};
