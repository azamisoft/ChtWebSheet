// Feature module; shared bindings are imported explicitly.
import {
  ArraySize,
  CellError,
  ErrorType,
  SimpleRangeValue,
  WEBSHEET_FORMULA_NA_SENTINEL,
  WEBSHEET_ROMAN_VALUES,
  X,
  loadExcelJsModule,
  loadJsZipModule,
  loadXlsxModule,
  state,
} from "../../app-core.js";
import {
  dataAnalysisCleanSignedZero,
  dataAnalysisComplexText,
  dataAnalysisParseComplexValue,
  dataAnalysisRegularizedBeta,
  dataAnalysisStudentTInverseTwoTail,
  dataAnalysisStudentTTwoTail,
} from "../data/forecast-and-statistics.js";
import {
  daysInMonthUtc,
  positiveModulo,
} from "../editing/autofill.js";
import {
  formulaArgumentBoolean,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  cellHasFormulaContent,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  currencySymbol,
  dateToExcelSerial,
  excelSerialToDate,
} from "../ui/help-and-utilities.js";
import {
  compareDateParts,
  datePartsForDatedif,
  datePartsToUtcDay,
} from "../workbook/saving-and-export.js";


async function loadWorkbookIoModules(options = {}) {
  const loaders = [loadExcelJsModule(), loadJsZipModule()];
  if (options.legacyXls) loaders.push(loadXlsxModule());
  await Promise.all(loaders);
}



function warmWorkbookIoModules() {
  if (typeof window !== "undefined" && window.__WEBSHEET_STANDALONE__) return;
  const warm = () => {
    void loadWorkbookIoModules();
  };
  if (typeof window !== "undefined" && typeof window.requestIdleCallback === "function") {
    window.requestIdleCallback(warm, { timeout: 5000 });
    return;
  }
  if (typeof window !== "undefined") window.setTimeout(warm, 1000);
}



function webSheetFormulaIsEmptyValue(value) {
  return value == null || value === "" || typeof value === "symbol";
}



function webSheetFormulaRawValue(value) {
  if (value && typeof value === "object" && typeof value.val === "number") return value.val;
  return value;
}



function webSheetFormulaFlattenValues(value) {
  if (value instanceof SimpleRangeValue) return value.valuesFromTopLeftCorner();
  if (Array.isArray(value)) return value.flatMap((row) => (Array.isArray(row) ? row : [row]));
  return [value];
}



function webSheetFormulaRangeData(value) {
  if (value instanceof SimpleRangeValue) return value.rawData();
  return [[value]];
}



function webSheetFormulaNumber(value) {
  if (value instanceof CellError) return value;
  const raw = webSheetFormulaRawValue(value);
  if (webSheetFormulaIsEmptyValue(raw)) return null;
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;
  if (typeof raw === "boolean") return raw ? 1 : 0;
  if (typeof raw === "string") {
    const trimmed = raw.trim();
    if (!trimmed) return null;
    const parsed = Number(trimmed.replace(/,/g, ""));
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}



function webSheetFormulaNumericValues(args) {
  const numbers = [];
  for (const arg of args) {
    for (const value of webSheetFormulaFlattenValues(arg)) {
      const number = webSheetFormulaNumber(value);
      if (number instanceof CellError) return number;
      if (Number.isFinite(number)) numbers.push(number);
    }
  }
  return numbers;
}



function webSheetFormulaAValues(args) {
  const values = [];
  for (const arg of args) {
    for (const value of webSheetFormulaFlattenValues(arg)) {
      if (value instanceof CellError) return value;
      const raw = webSheetFormulaRawValue(value);
      if (webSheetFormulaIsEmptyValue(raw)) continue;
      if (typeof raw === "number" && Number.isFinite(raw)) values.push(raw);
      else if (typeof raw === "boolean") values.push(raw ? 1 : 0);
      else if (typeof raw === "string") {
        const trimmed = raw.trim();
        const parsed = Number(trimmed.replace(/,/g, ""));
        values.push(trimmed && Number.isFinite(parsed) ? parsed : 0);
      } else {
        const number = Number(raw);
        values.push(Number.isFinite(number) ? number : 0);
      }
    }
  }
  return values;
}



function webSheetFormulaNumericOrZero(value) {
  const number = webSheetFormulaNumber(value);
  if (number instanceof CellError) return number;
  return Number.isFinite(number) ? number : 0;
}



function webSheetFormulaScalar(value) {
  if (value instanceof SimpleRangeValue) return value.valuesFromTopLeftCorner()[0];
  if (Array.isArray(value)) return Array.isArray(value[0]) ? value[0][0] : value[0];
  return value;
}



function webSheetFormulaText(value) {
  if (value instanceof CellError) return value;
  const raw = webSheetFormulaRawValue(value);
  if (webSheetFormulaIsEmptyValue(raw)) return "";
  if (typeof raw === "boolean") return raw ? "TRUE" : "FALSE";
  return String(raw);
}



function webSheetFormulaError(type) {
  return new CellError(type);
}



function webSheetFormulaCompareValues(left, right) {
  const leftNumber = webSheetFormulaNumber(left);
  const rightNumber = webSheetFormulaNumber(right);
  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) {
    return leftNumber === rightNumber ? 0 : leftNumber > rightNumber ? 1 : -1;
  }
  const leftText = webSheetFormulaText(left);
  const rightText = webSheetFormulaText(right);
  if (leftText instanceof CellError) return NaN;
  if (rightText instanceof CellError) return NaN;
  return leftText.localeCompare(rightText, undefined, { sensitivity: "accent", numeric: true });
}



function webSheetFormulaLetName(ast) {
  const name = String(ast?.expressionName || "").trim();
  return /^[A-Z_][A-Z0-9_.]*$/iu.test(name) ? name.toLocaleUpperCase() : "";
}



function webSheetFormulaLetAstContainsEnvName(ast, env) {
  if (!ast || typeof ast !== "object") return false;
  if (ast.type === "NAMED_EXPRESSION") return env.has(String(ast.expressionName || "").toLocaleUpperCase());
  if (Array.isArray(ast.args)) return ast.args.flat().some((item) => webSheetFormulaLetAstContainsEnvName(item, env));
  return Object.values(ast).some((value) => webSheetFormulaLetAstContainsEnvName(value, env));
}



function webSheetFormulaLetScalarNumber(value) {
  const number = webSheetFormulaNumber(value instanceof SimpleRangeValue ? value.valuesFromTopLeftCorner()[0] : value);
  return number instanceof CellError ? number : Number(number);
}



function webSheetFormulaLetLogical(value) {
  if (value instanceof CellError) return value;
  const raw = webSheetFormulaRawValue(value instanceof SimpleRangeValue ? value.valuesFromTopLeftCorner()[0] : value);
  if (raw instanceof CellError) return raw;
  if (typeof raw === "boolean") return raw;
  const number = webSheetFormulaNumber(raw);
  if (Number.isFinite(number)) return number !== 0;
  const text = String(raw ?? "").trim().toLocaleUpperCase();
  if (text === "TRUE") return true;
  if (text === "FALSE") return false;
  return Boolean(raw);
}



function webSheetFormulaLetApplyBinary(ast, state, plugin, env, operation) {
  const left = webSheetFormulaLetAstValue(ast.left, state, plugin, env);
  if (left instanceof CellError) return left;
  const right = webSheetFormulaLetAstValue(ast.right, state, plugin, env);
  if (right instanceof CellError) return right;
  if (operation === "CONCATENATE_OP") {
    const leftText = webSheetFormulaText(left);
    const rightText = webSheetFormulaText(right);
    if (leftText instanceof CellError) return leftText;
    if (rightText instanceof CellError) return rightText;
    return `${leftText}${rightText}`;
  }
  if (/^(EQUALS|NOT_EQUAL|GREATER_THAN|LESS_THAN|GREATER_THAN_OR_EQUAL|LESS_THAN_OR_EQUAL)_OP$/u.test(operation)) {
    const comparison = webSheetFormulaCompareValues(left, right);
    if (!Number.isFinite(comparison)) return webSheetFormulaError(ErrorType.VALUE);
    if (operation === "EQUALS_OP") return comparison === 0;
    if (operation === "NOT_EQUAL_OP") return comparison !== 0;
    if (operation === "GREATER_THAN_OP") return comparison > 0;
    if (operation === "LESS_THAN_OP") return comparison < 0;
    if (operation === "GREATER_THAN_OR_EQUAL_OP") return comparison >= 0;
    if (operation === "LESS_THAN_OR_EQUAL_OP") return comparison <= 0;
  }
  const leftNumber = webSheetFormulaLetScalarNumber(left);
  if (!Number.isFinite(leftNumber)) return leftNumber instanceof CellError ? leftNumber : webSheetFormulaError(ErrorType.VALUE);
  const rightNumber = webSheetFormulaLetScalarNumber(right);
  if (!Number.isFinite(rightNumber)) return rightNumber instanceof CellError ? rightNumber : webSheetFormulaError(ErrorType.VALUE);
  if (operation === "PLUS_OP") return leftNumber + rightNumber;
  if (operation === "MINUS_OP") return leftNumber - rightNumber;
  if (operation === "TIMES_OP") return leftNumber * rightNumber;
  if (operation === "DIV_OP") return rightNumber === 0 ? webSheetFormulaError(ErrorType.DIV_BY_ZERO) : leftNumber / rightNumber;
  if (operation === "POWER_OP") return leftNumber ** rightNumber;
  return webSheetFormulaError(ErrorType.VALUE);
}



function webSheetFormulaLetFunctionValue(name, args) {
  const functionName = String(name || "").toLocaleUpperCase();
  if (args.some((value) => value instanceof CellError)) return args.find((value) => value instanceof CellError);
  if (functionName === "SUM") {
    const values = webSheetFormulaNumericValues(args);
    return values instanceof CellError ? values : values.reduce((sum, value) => sum + value, 0);
  }
  if (functionName === "AVERAGE") {
    const values = webSheetFormulaNumericValues(args);
    return values instanceof CellError ? values : values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  }
  if (functionName === "COUNT") {
    const values = webSheetFormulaNumericValues(args);
    return values instanceof CellError ? values : values.length;
  }
  if (functionName === "COUNTA") {
    return args.flatMap((item) => webSheetFormulaFlattenValues(item)).filter((value) => !webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(value))).length;
  }
  if (functionName === "MAX" || functionName === "MIN") {
    const values = webSheetFormulaNumericValues(args);
    if (values instanceof CellError) return values;
    if (!values.length) return 0;
    return functionName === "MAX" ? Math.max(...values) : Math.min(...values);
  }
  if (functionName === "PRODUCT") {
    const values = webSheetFormulaNumericValues(args);
    return values instanceof CellError ? values : values.reduce((product, value) => product * value, 1);
  }
  if (functionName === "IF") return webSheetFormulaLetLogical(args[0]) ? args[1] : args[2];
  if (functionName === "IFS") {
    for (let index = 0; index < args.length; index += 2) {
      if (webSheetFormulaLetLogical(args[index])) return args[index + 1];
    }
    return webSheetFormulaError(ErrorType.NA);
  }
  if (functionName === "AND") return args.every((value) => webSheetFormulaLetLogical(value));
  if (functionName === "OR") return args.some((value) => webSheetFormulaLetLogical(value));
  if (functionName === "XOR") return args.filter((value) => webSheetFormulaLetLogical(value)).length % 2 === 1;
  if (functionName === "NOT") return !webSheetFormulaLetLogical(args[0]);
  if (functionName === "CONCAT") return args.flatMap((item) => webSheetFormulaFlattenValues(item)).map((value) => webSheetFormulaText(value)).join("");
  if (functionName === "TEXTJOIN") {
    const delimiter = webSheetFormulaText(args[0]);
    if (delimiter instanceof CellError) return delimiter;
    const ignoreEmpty = webSheetFormulaLetLogical(args[1]);
    const parts = [];
    args.slice(2).forEach((item) => {
      webSheetFormulaFlattenValues(item).forEach((value) => {
        const text = webSheetFormulaText(value);
        if (!(ignoreEmpty && text === "")) parts.push(text);
      });
    });
    return parts.find((item) => item instanceof CellError) || parts.join(delimiter);
  }
  if (functionName === "AGGREGATE") return webSheetFormulaAggregate(args[0], args[1], args.slice(2));
  return null;
}



function webSheetFormulaLetAstValue(ast, state, plugin, env) {
  if (!ast || typeof ast !== "object") return null;
  switch (ast.type) {
    case "EMPTY":
      return null;
    case "NUMBER":
    case "STRING":
      return ast.value;
    case "NAMED_EXPRESSION": {
      const name = String(ast.expressionName || "").toLocaleUpperCase();
      if (name === "TRUE") return true;
      if (name === "FALSE") return false;
      if (env.has(name)) return env.get(name);
      return plugin.evaluateAst(ast, state);
    }
    case "CELL_REFERENCE":
    case "CELL_RANGE":
    case "COLUMN_RANGE":
    case "ROW_RANGE":
      return plugin.evaluateAst(ast, state);
    case "PARENTHESIS":
      return webSheetFormulaLetAstValue(ast.expression, state, plugin, env);
    case "PLUS_UNARY_OP":
      return webSheetFormulaLetScalarNumber(webSheetFormulaLetAstValue(ast.value, state, plugin, env));
    case "MINUS_UNARY_OP": {
      const value = webSheetFormulaLetScalarNumber(webSheetFormulaLetAstValue(ast.value, state, plugin, env));
      return value instanceof CellError ? value : -value;
    }
    case "PERCENT_OP": {
      const value = webSheetFormulaLetScalarNumber(webSheetFormulaLetAstValue(ast.value, state, plugin, env));
      return value instanceof CellError ? value : value / 100;
    }
    case "ARRAY":
      return SimpleRangeValue.onlyValues((ast.args || []).map((row) => row.map((item) => webSheetFormulaLetAstValue(item, state, plugin, env))));
    case "FUNCTION_CALL": {
      if (!webSheetFormulaLetAstContainsEnvName(ast, env)) return plugin.evaluateAst(ast, state);
      const args = (ast.args || []).map((arg) => webSheetFormulaLetAstValue(arg, state, plugin, env));
      const value = webSheetFormulaLetFunctionValue(ast.procedureName, args);
      return value == null ? webSheetFormulaError(ErrorType.NAME) : value;
    }
    case "PLUS_OP":
    case "MINUS_OP":
    case "TIMES_OP":
    case "DIV_OP":
    case "POWER_OP":
    case "CONCATENATE_OP":
    case "EQUALS_OP":
    case "NOT_EQUAL_OP":
    case "GREATER_THAN_OP":
    case "LESS_THAN_OP":
    case "GREATER_THAN_OR_EQUAL_OP":
    case "LESS_THAN_OR_EQUAL_OP":
      return webSheetFormulaLetApplyBinary(ast, state, plugin, env, ast.type);
    default:
      return plugin.evaluateAst(ast, state);
  }
}



function webSheetFormulaLambdaSpec(lambdaAst) {
  if (!lambdaAst || lambdaAst.type !== "FUNCTION_CALL" || String(lambdaAst.procedureName || "").toLocaleUpperCase() !== "LAMBDA") {
    return null;
  }
  const args = lambdaAst.args || [];
  if (args.length < 2) return null;
  const names = args.slice(0, -1).map(webSheetFormulaLetName);
  if (names.some((name) => !name)) return null;
  return { names, body: args[args.length - 1] };
}



function webSheetFormulaApplyLambda(lambdaAst, values, state, plugin) {
  const spec = webSheetFormulaLambdaSpec(lambdaAst);
  if (!spec || values.length !== spec.names.length) return webSheetFormulaError(ErrorType.VALUE);
  const env = new Map();
  spec.names.forEach((name, index) => env.set(name, values[index]));
  return webSheetFormulaLetAstValue(spec.body, state, plugin, env);
}



function webSheetFormulaMakeArraySize(ast) {
  return new ArraySize(
    Math.max(1, Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[1], 1))),
    Math.max(1, Math.trunc(webSheetFormulaAstNumberLiteral(ast.args?.[0], 1))),
  );
}



function webSheetFormulaByRowArraySize(ast, state, plugin) {
  const base = plugin.arraySizeForAst(ast.args?.[0], state);
  return new ArraySize(1, Math.max(1, base?.height || 1));
}



function webSheetFormulaByColArraySize(ast, state, plugin) {
  const base = plugin.arraySizeForAst(ast.args?.[0], state);
  return new ArraySize(Math.max(1, base?.width || 1), 1);
}



function webSheetFormulaWildcardRegex(pattern) {
  const source = [...String(pattern)].map((char) => {
    if (char === "*") return ".*";
    if (char === "?") return ".";
    return char.replace(/[.+^${}()|[\]\\]/g, "\\$&");
  }).join("");
  return new RegExp(`^${source}$`, "iu");
}



function webSheetFormulaCriteriaMatcher(criteria) {
  const rawCriteria = webSheetFormulaRawValue(criteria);
  if (rawCriteria instanceof CellError) return () => false;
  const text = String(rawCriteria ?? "").trim();
  const match = text.match(/^(<=|>=|<>|=|<|>)(.*)$/u);
  const operator = match ? match[1] : "=";
  const targetText = match ? match[2].trim() : text;
  const hasWildcard = /(?<!~)[*?]/u.test(targetText);
  const regex = hasWildcard ? webSheetFormulaWildcardRegex(targetText.replace(/~([*?])/gu, "$1")) : null;
  const numericTarget = targetText === "" ? null : Number(targetText.replace(/,/g, ""));
  return (value) => {
    if (value instanceof CellError) return false;
    if (regex) {
      const actual = webSheetFormulaText(value);
      if (actual instanceof CellError) return false;
      const matched = regex.test(actual);
      return operator === "<>" ? !matched : matched;
    }
    const target = Number.isFinite(numericTarget) ? numericTarget : targetText;
    const comparison = webSheetFormulaCompareValues(value, target);
    if (!Number.isFinite(comparison)) return false;
    if (operator === "=") return comparison === 0;
    if (operator === "<>") return comparison !== 0;
    if (operator === ">") return comparison > 0;
    if (operator === ">=") return comparison >= 0;
    if (operator === "<") return comparison < 0;
    if (operator === "<=") return comparison <= 0;
    return false;
  };
}



function webSheetFormulaCriteriaPairs(targetData, criteriaArgs) {
  const pairs = [];
  for (let index = 0; index < criteriaArgs.length; index += 2) {
    const data = webSheetFormulaRangeData(criteriaArgs[index]);
    if (data.length !== targetData.length || data.some((row, rowIndex) => row.length !== (targetData[rowIndex] || []).length)) {
      return webSheetFormulaError(ErrorType.VALUE);
    }
    pairs.push({ data, matches: webSheetFormulaCriteriaMatcher(criteriaArgs[index + 1]) });
  }
  return pairs;
}



function webSheetFormulaConditionalValues(targetRange, criteriaArgs, { numericOnly = true } = {}) {
  const targetData = webSheetFormulaRangeData(targetRange);
  const pairs = webSheetFormulaCriteriaPairs(targetData, criteriaArgs);
  if (pairs instanceof CellError) return pairs;
  const values = [];
  targetData.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      if (!pairs.every((pair) => pair.matches(pair.data[rowIndex][colIndex]))) return;
      if (!numericOnly) {
        values.push(value);
        return;
      }
      const number = webSheetFormulaNumber(value);
      if (number instanceof CellError) values.push(number);
      else if (Number.isFinite(number)) values.push(number);
    });
  });
  const error = values.find((value) => value instanceof CellError);
  return error || values;
}



function webSheetFormulaSumProduct(args) {
  const matrices = args.map((arg) => webSheetFormulaMatrix(arg));
  const height = Math.max(1, ...matrices.map((matrix) => matrix.length));
  const width = Math.max(1, ...matrices.flatMap((matrix) => matrix.map((row) => row.length)));
  const canBroadcast = (matrix) => matrix.length === 1 && (matrix[0]?.length || 1) === 1;
  if (matrices.some((matrix) => !canBroadcast(matrix) && (matrix.length !== height || matrix.some((row) => row.length !== width)))) {
    return webSheetFormulaError(ErrorType.VALUE);
  }
  let total = 0;
  for (let row = 0; row < height; row += 1) {
    for (let col = 0; col < width; col += 1) {
      let product = 1;
      for (const matrix of matrices) {
        const value = canBroadcast(matrix) ? matrix[0]?.[0] : matrix[row]?.[col];
        const number = webSheetFormulaNumericOrZero(value);
        if (number instanceof CellError) return number;
        product *= number;
      }
      total += product;
    }
  }
  return total;
}



function webSheetFormulaCountBlank(range) {
  return webSheetFormulaFlattenValues(range)
    .filter((value) => webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(value)))
    .length;
}



function webSheetFormulaSumIf(range, criteria, sumRange = range) {
  const criteriaData = webSheetFormulaRangeData(range);
  const sumData = webSheetFormulaRangeData(sumRange);
  const matcher = webSheetFormulaCriteriaMatcher(criteria);
  let total = 0;
  criteriaData.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      if (!matcher(value)) return;
      const number = webSheetFormulaNumericOrZero(sumData[rowIndex]?.[colIndex]);
      if (number instanceof CellError) return;
      total += number;
    });
  });
  return total;
}



function webSheetFormulaAverageIf(range, criteria, averageRange = range) {
  const criteriaData = webSheetFormulaRangeData(range);
  const averageData = webSheetFormulaRangeData(averageRange);
  const matcher = webSheetFormulaCriteriaMatcher(criteria);
  const matched = [];
  criteriaData.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      if (!matcher(value)) return;
      const number = webSheetFormulaNumber(averageData[rowIndex]?.[colIndex]);
      if (number instanceof CellError) matched.push(number);
      else if (Number.isFinite(number)) matched.push(number);
    });
  });
  const error = matched.find((value) => value instanceof CellError);
  if (error) return error;
  if (!matched.length) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  return matched.reduce((sum, value) => sum + value, 0) / matched.length;
}



function webSheetFormulaDateSerial(value) {
  const raw = webSheetFormulaRawValue(webSheetFormulaScalar(value));
  if (raw instanceof CellError) return raw;
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) return dateToExcelSerial(raw);
  const number = typeof raw === "number" ? raw : Number(String(raw ?? "").trim());
  return Number.isFinite(number) ? Math.trunc(number) : webSheetFormulaError(ErrorType.VALUE);
}



function webSheetFormulaDateParts(value) {
  const raw = webSheetFormulaRawValue(webSheetFormulaScalar(value));
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return datePartsForDatedif(raw.getFullYear(), raw.getMonth() + 1, raw.getDate());
  }
  const serial = webSheetFormulaDateSerial(raw);
  if (serial instanceof CellError) return serial;
  const date = excelSerialToDate(serial);
  return datePartsForDatedif(date.getFullYear(), date.getMonth() + 1, date.getDate());
}



function webSheetFormulaDatePartsToSerial(parts) {
  return Math.floor(Date.UTC(parts.year, parts.month - 1, parts.day) / 86400000) + 25569;
}



function webSheetFormulaValidDateParts(year, month, day) {
  const y = Number(year);
  const m = Number(month);
  const d = Number(day);
  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d) || y < 1 || m < 1 || m > 12 || d < 1) return null;
  if (d > daysInMonthUtc(y, m - 1)) return null;
  return datePartsForDatedif(y, m, d);
}



function webSheetFormulaParseDateText(value) {
  const raw = webSheetFormulaRawValue(webSheetFormulaScalar(value));
  if (raw instanceof CellError) return raw;
  const text = String(raw ?? "").trim();
  if (!text) return webSheetFormulaError(ErrorType.VALUE);
  let match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/u);
  if (match) return webSheetFormulaValidDateParts(match[1], match[2], match[3]) || webSheetFormulaError(ErrorType.VALUE);
  match = text.match(/^(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日$/u);
  if (match) return webSheetFormulaValidDateParts(match[1], match[2], match[3]) || webSheetFormulaError(ErrorType.VALUE);
  match = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/u);
  if (match) return webSheetFormulaValidDateParts(match[3], match[1], match[2]) || webSheetFormulaError(ErrorType.VALUE);
  const parsed = Date.parse(text);
  if (Number.isFinite(parsed)) {
    const date = new Date(parsed);
    return webSheetFormulaValidDateParts(date.getFullYear(), date.getMonth() + 1, date.getDate()) || webSheetFormulaError(ErrorType.VALUE);
  }
  return webSheetFormulaError(ErrorType.VALUE);
}



function webSheetFormulaDateValue(value) {
  const parts = webSheetFormulaParseDateText(value);
  return parts instanceof CellError ? parts : webSheetFormulaDatePartsToSerial(parts);
}



function webSheetFormulaParseTimeText(value) {
  const raw = webSheetFormulaRawValue(webSheetFormulaScalar(value));
  if (raw instanceof CellError) return raw;
  let text = String(raw ?? "").trim();
  if (!text) return webSheetFormulaError(ErrorType.VALUE);
  const japanesePm = /^午後/u.test(text);
  const japaneseAm = /^午前/u.test(text);
  text = text.replace(/^午前|^午後/u, "").trim();
  const ampmMatch = text.match(/\s*(AM|PM|A|P)$/iu);
  const ampm = ampmMatch?.[1]?.toUpperCase();
  if (ampmMatch) text = text.slice(0, ampmMatch.index).trim();
  const timePart = text.includes(" ") ? text.split(/\s+/u).at(-1) : text;
  const match = timePart.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?$/u);
  if (!match) return webSheetFormulaError(ErrorType.VALUE);
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = Number(match[3] ?? 0);
  if (!Number.isFinite(hour) || !Number.isFinite(minute) || !Number.isFinite(second)) return webSheetFormulaError(ErrorType.VALUE);
  if (minute < 0 || minute >= 60 || second < 0 || second >= 60) return webSheetFormulaError(ErrorType.VALUE);
  const hasAmPm = Boolean(ampm || japaneseAm || japanesePm);
  if (hasAmPm) {
    if (hour < 1 || hour > 12) return webSheetFormulaError(ErrorType.VALUE);
    if (ampm === "PM" || ampm === "P" || japanesePm) hour = hour % 12 + 12;
    else hour %= 12;
  } else if (hour < 0 || hour >= 24) {
    return webSheetFormulaError(ErrorType.VALUE);
  }
  return (hour * 3600 + minute * 60 + second) / 86400;
}



function webSheetFormulaTimeValue(value) {
  return webSheetFormulaParseTimeText(value);
}



function webSheetFormulaTimeSerial(value) {
  const raw = webSheetFormulaRawValue(webSheetFormulaScalar(value));
  if (raw instanceof CellError) return raw;
  const number = typeof raw === "number" ? raw : webSheetFormulaParseTimeText(raw);
  if (number instanceof CellError) return number;
  if (!Number.isFinite(number) || number < 0) return webSheetFormulaError(ErrorType.NUM);
  return positiveModulo(number, 1);
}



function webSheetFormulaTimeUnit(value, unit) {
  const fraction = webSheetFormulaTimeSerial(value);
  if (fraction instanceof CellError) return fraction;
  const totalSeconds = Math.floor(fraction * 86400 + 1e-9) % 86400;
  if (unit === "hour") return Math.floor(totalSeconds / 3600);
  if (unit === "minute") return Math.floor((totalSeconds % 3600) / 60);
  return totalSeconds % 60;
}



function webSheetFormulaAddMonths(startDate, months, endOfMonth = false) {
  const start = webSheetFormulaDateParts(startDate);
  if (start instanceof CellError) return start;
  const amount = Math.trunc(Number(months) || 0);
  const targetMonthIndex = start.month - 1 + amount;
  const targetYear = start.year + Math.floor(targetMonthIndex / 12);
  const targetMonth = positiveModulo(targetMonthIndex, 12) + 1;
  const targetDay = endOfMonth ? daysInMonthUtc(targetYear, targetMonth - 1) : Math.min(start.day, daysInMonthUtc(targetYear, targetMonth - 1));
  return webSheetFormulaDatePartsToSerial(datePartsForDatedif(targetYear, targetMonth, targetDay));
}



function webSheetFormulaDays360(startDate, endDate, method = false) {
  const start = webSheetFormulaDateParts(startDate);
  const end = webSheetFormulaDateParts(endDate);
  if (start instanceof CellError) return start;
  if (end instanceof CellError) return end;
  let startDay = start.day;
  let endDay = end.day;
  const european = Boolean(method);
  if (european) {
    startDay = Math.min(startDay, 30);
    endDay = Math.min(endDay, 30);
  } else {
    const startLastDay = start.day === daysInMonthUtc(start.year, start.month - 1);
    const endLastDay = end.day === daysInMonthUtc(end.year, end.month - 1);
    if (startLastDay) startDay = 30;
    if (endLastDay && startDay >= 30) endDay = 30;
  }
  return (end.year - start.year) * 360 + (end.month - start.month) * 30 + (endDay - startDay);
}



function webSheetFormulaYearFrac(startDate, endDate, basis = 0) {
  const start = webSheetFormulaDateParts(startDate);
  const end = webSheetFormulaDateParts(endDate);
  if (start instanceof CellError) return start;
  if (end instanceof CellError) return end;
  if (compareDateParts(start, end) > 0) return webSheetFormulaYearFrac(endDate, startDate, basis);
  const mode = Math.trunc(Number(basis) || 0);
  const days = datePartsToUtcDay(end) - datePartsToUtcDay(start);
  if (mode === 0) return webSheetFormulaDays360(startDate, endDate, false) / 360;
  if (mode === 2) return days / 360;
  if (mode === 3) return days / 365;
  if (mode === 4) return webSheetFormulaDays360(startDate, endDate, true) / 360;
  if (start.year === end.year) return days / daysInYearUtc(start.year);
  let total = 0;
  total += (datePartsToUtcDay(datePartsForDatedif(start.year + 1, 1, 1)) - datePartsToUtcDay(start)) / daysInYearUtc(start.year);
  for (let year = start.year + 1; year < end.year; year += 1) total += 1;
  total += (datePartsToUtcDay(end) - datePartsToUtcDay(datePartsForDatedif(end.year, 1, 1))) / daysInYearUtc(end.year);
  return total;
}



function daysInYearUtc(year) {
  return daysInMonthUtc(year, 1) === 29 ? 366 : 365;
}



function webSheetFormulaWeekendMask(weekend) {
  const raw = webSheetFormulaRawValue(webSheetFormulaScalar(weekend));
  if (raw == null || raw === "" || raw === WEBSHEET_FORMULA_NA_SENTINEL) return new Set([0, 6]);
  const text = String(raw).trim();
  if (/^[01]{7}$/u.test(text)) {
    const mask = new Set();
    [...text].forEach((char, index) => {
      if (char === "1") mask.add(index === 6 ? 0 : index + 1);
    });
    return mask;
  }
  const code = Math.trunc(Number(raw));
  const weekendCodes = {
    1: [0, 6],
    2: [0, 1],
    3: [1, 2],
    4: [2, 3],
    5: [3, 4],
    6: [4, 5],
    7: [5, 6],
    11: [0],
    12: [1],
    13: [2],
    14: [3],
    15: [4],
    16: [5],
    17: [6],
  };
  return weekendCodes[code] ? new Set(weekendCodes[code]) : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaHolidaySet(holidays) {
  if (holidays == null || holidays === WEBSHEET_FORMULA_NA_SENTINEL) return new Set();
  const set = new Set();
  for (const value of webSheetFormulaFlattenValues(holidays)) {
    if (webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(value))) continue;
    const serial = webSheetFormulaDateSerial(value);
    if (serial instanceof CellError) return serial;
    set.add(serial);
  }
  return set;
}



function webSheetFormulaDateSerialWeekday(serial) {
  const value = Math.trunc(Number(serial));
  if (!Number.isFinite(value)) return NaN;
  return ((value + 6) % 7 + 7) % 7;
}



function webSheetFormulaWeekStartForReturnType(returnType, defaultType = 1) {
  const type = Math.trunc(Number(returnType == null || returnType === "" ? defaultType : returnType));
  if (type === 1) return 0;
  if (type === 2 || type === 11) return 1;
  if (type >= 12 && type <= 16) return type - 10;
  if (type === 17) return 0;
  return null;
}



function webSheetFormulaWeekdayValue(serialNumber, returnType = 1) {
  const serial = webSheetFormulaDateSerial(serialNumber);
  if (serial instanceof CellError) return serial;
  const day = webSheetFormulaDateSerialWeekday(serial);
  const type = Math.trunc(Number(returnType == null || returnType === "" ? 1 : returnType));
  if (type === 3) return positiveModulo(day - 1, 7);
  const weekStart = webSheetFormulaWeekStartForReturnType(type, 1);
  return weekStart == null ? webSheetFormulaError(ErrorType.NUM) : positiveModulo(day - weekStart, 7) + 1;
}



function webSheetFormulaWeeknumValue(serialNumber, returnType = 1) {
  const serial = webSheetFormulaDateSerial(serialNumber);
  if (serial instanceof CellError) return serial;
  const type = Math.trunc(Number(returnType == null || returnType === "" ? 1 : returnType));
  if (type === 21) return webSheetFormulaIsoWeeknumValue(serial);
  const weekStart = webSheetFormulaWeekStartForReturnType(type, 1);
  if (weekStart == null || type === 3) return webSheetFormulaError(ErrorType.NUM);
  const parts = webSheetFormulaDateParts(serial);
  if (parts instanceof CellError) return parts;
  const jan1Serial = webSheetFormulaDatePartsToSerial(datePartsForDatedif(parts.year, 1, 1));
  const jan1Day = webSheetFormulaDateSerialWeekday(jan1Serial);
  const weekOneStart = jan1Serial - positiveModulo(jan1Day - weekStart, 7);
  return Math.floor((Math.trunc(serial) - weekOneStart) / 7) + 1;
}



function webSheetFormulaIsoWeeknumValue(serialNumber) {
  const parts = typeof serialNumber === "number" ? webSheetFormulaDateParts(serialNumber) : webSheetFormulaDateParts(serialNumber);
  if (parts instanceof CellError) return parts;
  const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
}



function webSheetFormulaIsWorkday(serial, weekendMask, holidays) {
  const day = webSheetFormulaDateSerialWeekday(serial);
  return !weekendMask.has(day) && !holidays.has(Math.trunc(serial));
}



function webSheetFormulaNetworkDays(startDate, endDate, weekend, holidays) {
  const start = webSheetFormulaDateSerial(startDate);
  const end = webSheetFormulaDateSerial(endDate);
  if (start instanceof CellError) return start;
  if (end instanceof CellError) return end;
  const weekendMask = webSheetFormulaWeekendMask(weekend);
  if (weekendMask instanceof CellError) return weekendMask;
  const holidaySet = webSheetFormulaHolidaySet(holidays);
  if (holidaySet instanceof CellError) return holidaySet;
  const direction = start <= end ? 1 : -1;
  let count = 0;
  for (let serial = start; direction > 0 ? serial <= end : serial >= end; serial += direction) {
    if (webSheetFormulaIsWorkday(serial, weekendMask, holidaySet)) count += direction;
  }
  return count;
}



function webSheetFormulaWorkday(startDate, days, weekend, holidays) {
  const start = webSheetFormulaDateSerial(startDate);
  if (start instanceof CellError) return start;
  const amount = Math.trunc(Number(webSheetFormulaRawValue(days)) || 0);
  const weekendMask = webSheetFormulaWeekendMask(weekend);
  if (weekendMask instanceof CellError) return weekendMask;
  const holidaySet = webSheetFormulaHolidaySet(holidays);
  if (holidaySet instanceof CellError) return holidaySet;
  if (!amount) return start;
  const direction = amount > 0 ? 1 : -1;
  let remaining = Math.abs(amount);
  let serial = start;
  while (remaining > 0) {
    serial += direction;
    if (webSheetFormulaIsWorkday(serial, weekendMask, holidaySet)) remaining -= 1;
  }
  return serial;
}



function webSheetFormulaCeilingMath(number, significance = 1, mode = 0) {
  const value = Number(number);
  const step = Math.abs(Number(significance || 1));
  if (!Number.isFinite(value) || !Number.isFinite(step) || !step) return webSheetFormulaError(ErrorType.NUM);
  if (!value) return 0;
  if (value > 0) return Math.ceil(value / step) * step;
  return mode ? Math.floor(value / step) * step : Math.ceil(value / step) * step;
}



function webSheetFormulaFloorMath(number, significance = 1, mode = 0) {
  const value = Number(number);
  const step = Math.abs(Number(significance || 1));
  if (!Number.isFinite(value) || !Number.isFinite(step) || !step) return webSheetFormulaError(ErrorType.NUM);
  if (!value) return 0;
  if (value > 0) return Math.floor(value / step) * step;
  return mode ? Math.ceil(value / step) * step : Math.floor(value / step) * step;
}



function webSheetFormulaMround(number, multiple) {
  const value = Number(number);
  const step = Number(multiple);
  if (!Number.isFinite(value) || !Number.isFinite(step) || !step || Math.sign(value) !== Math.sign(step)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  return Math.sign(value) * Math.round(Math.abs(value / step)) * Math.abs(step);
}



function webSheetFormulaCeilingPrecise(number, significance = 1) {
  const value = Number(number);
  const step = Math.abs(Number(significance == null || significance === "" ? 1 : significance));
  if (!Number.isFinite(value) || !Number.isFinite(step)) return webSheetFormulaError(ErrorType.NUM);
  if (!step || !value) return 0;
  return Math.ceil(value / step) * step;
}



function webSheetFormulaFloorPrecise(number, significance = 1) {
  const value = Number(number);
  const step = Math.abs(Number(significance == null || significance === "" ? 1 : significance));
  if (!Number.isFinite(value) || !Number.isFinite(step)) return webSheetFormulaError(ErrorType.NUM);
  if (!step || !value) return 0;
  return Math.floor(value / step) * step;
}



function webSheetFormulaRoundAwayToParity(number, parity) {
  const value = Number(number);
  if (!Number.isFinite(value)) return webSheetFormulaError(ErrorType.NUM);
  if (value === 0) return parity ? 1 : 0;
  const sign = value < 0 ? -1 : 1;
  let rounded = Math.ceil(Math.abs(value));
  while (rounded % 2 !== parity) rounded += 1;
  return sign * rounded;
}



function webSheetFormulaSeriesSum(x, n, m, coefficients) {
  const values = webSheetFormulaNumericValues([coefficients]);
  if (values instanceof CellError) return values;
  const base = Number(x);
  const startPower = Number(n);
  const step = Number(m);
  if (!Number.isFinite(base) || !Number.isFinite(startPower) || !Number.isFinite(step)) return webSheetFormulaError(ErrorType.NUM);
  return values.reduce((sum, coefficient, index) => sum + coefficient * (base ** (startPower + index * step)), 0);
}



function webSheetFormulaFinancialType(type = 0) {
  const value = Math.trunc(Number(type) || 0);
  return value === 0 || value === 1 ? value : null;
}



function webSheetFormulaPayment(rate, nper, pv, fv = 0, type = 0) {
  const r = Number(rate);
  const periods = Number(nper);
  const present = Number(pv);
  const future = Number(fv || 0);
  const paymentType = webSheetFormulaFinancialType(type);
  if (![r, periods, present, future].every(Number.isFinite) || paymentType == null || periods === 0) return webSheetFormulaError(ErrorType.NUM);
  if (r === 0) return -(present + future) / periods;
  const factor = (1 + r) ** periods;
  return -(r * (future + present * factor)) / ((1 + r * paymentType) * (factor - 1));
}



function webSheetFormulaFutureValue(rate, nper, pmt, pv = 0, type = 0) {
  const r = Number(rate);
  const periods = Number(nper);
  const payment = Number(pmt);
  const present = Number(pv || 0);
  const paymentType = webSheetFormulaFinancialType(type);
  if (![r, periods, payment, present].every(Number.isFinite) || paymentType == null) return webSheetFormulaError(ErrorType.NUM);
  if (r === 0) return -(present + payment * periods);
  const factor = (1 + r) ** periods;
  return -(present * factor + payment * (1 + r * paymentType) * (factor - 1) / r);
}



function webSheetFormulaPresentValue(rate, nper, pmt, fv = 0, type = 0) {
  const r = Number(rate);
  const periods = Number(nper);
  const payment = Number(pmt);
  const future = Number(fv || 0);
  const paymentType = webSheetFormulaFinancialType(type);
  if (![r, periods, payment, future].every(Number.isFinite) || paymentType == null) return webSheetFormulaError(ErrorType.NUM);
  if (r === 0) return -(future + payment * periods);
  const factor = (1 + r) ** periods;
  return -(future + payment * (1 + r * paymentType) * (factor - 1) / r) / factor;
}



function webSheetFormulaNper(rate, pmt, pv, fv = 0, type = 0) {
  const r = Number(rate);
  const payment = Number(pmt);
  const present = Number(pv);
  const future = Number(fv || 0);
  const paymentType = webSheetFormulaFinancialType(type);
  if (![r, payment, present, future].every(Number.isFinite) || paymentType == null) return webSheetFormulaError(ErrorType.NUM);
  if (r === 0) return payment === 0 ? webSheetFormulaError(ErrorType.DIV_BY_ZERO) : -(present + future) / payment;
  const adjustedPayment = payment * (1 + r * paymentType);
  const numerator = adjustedPayment - future * r;
  const denominator = present * r + adjustedPayment;
  const ratio = numerator / denominator;
  if (!Number.isFinite(ratio) || ratio <= 0 || 1 + r <= 0) return webSheetFormulaError(ErrorType.NUM);
  return Math.log(ratio) / Math.log(1 + r);
}



function webSheetFormulaRate(nper, pmt, pv, fv = 0, type = 0, guess = 0.1) {
  const periods = Number(nper);
  const payment = Number(pmt);
  const present = Number(pv);
  const future = Number(fv || 0);
  const paymentType = webSheetFormulaFinancialType(type);
  if (![periods, payment, present, future].every(Number.isFinite) || paymentType == null || periods <= 0) return webSheetFormulaError(ErrorType.NUM);
  const valueAt = (rate) => {
    if (Math.abs(rate) < 1e-10) return present + payment * periods + future;
    const factor = (1 + rate) ** periods;
    return present * factor + payment * (1 + rate * paymentType) * (factor - 1) / rate + future;
  };
  let rate = Number.isFinite(Number(guess)) ? Number(guess) : 0.1;
  for (let index = 0; index < 80; index += 1) {
    const f = valueAt(rate);
    const delta = 1e-6;
    const derivative = (valueAt(rate + delta) - f) / delta;
    if (!Number.isFinite(f) || !Number.isFinite(derivative) || Math.abs(derivative) < 1e-12) break;
    const next = rate - f / derivative;
    if (!Number.isFinite(next) || next <= -0.999999999) break;
    if (Math.abs(next - rate) < 1e-10) return next;
    rate = next;
  }
  let low = -0.999999;
  let high = Math.max(1, rate || 1);
  let fLow = valueAt(low);
  let fHigh = valueAt(high);
  for (let index = 0; index < 80 && Math.sign(fLow) === Math.sign(fHigh); index += 1) {
    high *= 2;
    fHigh = valueAt(high);
  }
  if (!Number.isFinite(fLow) || !Number.isFinite(fHigh) || Math.sign(fLow) === Math.sign(fHigh)) return webSheetFormulaError(ErrorType.NUM);
  for (let index = 0; index < 100; index += 1) {
    const mid = (low + high) / 2;
    const fMid = valueAt(mid);
    if (!Number.isFinite(fMid)) return webSheetFormulaError(ErrorType.NUM);
    if (Math.abs(fMid) < 1e-8 || Math.abs(high - low) < 1e-10) return mid;
    if (Math.sign(fLow) === Math.sign(fMid)) {
      low = mid;
      fLow = fMid;
    } else {
      high = mid;
    }
  }
  return webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaFinancialValues(values) {
  const numbers = [];
  for (const value of webSheetFormulaFlattenValues(values)) {
    const number = webSheetFormulaNumber(value);
    if (number instanceof CellError) return number;
    if (Number.isFinite(number)) numbers.push(number);
  }
  return numbers;
}



function webSheetFormulaNpv(rate, values) {
  const r = Number(rate);
  if (!Number.isFinite(r) || r === -1) return webSheetFormulaError(ErrorType.NUM);
  let period = 1;
  let total = 0;
  for (const arg of values) {
    const numbers = webSheetFormulaFinancialValues(arg);
    if (numbers instanceof CellError) return numbers;
    for (const number of numbers) {
      total += number / ((1 + r) ** period);
      period += 1;
    }
  }
  return total;
}



function webSheetFormulaFvSchedule(principal, schedule) {
  const base = Number(principal);
  const rates = webSheetFormulaNumericValues([schedule]);
  if (rates instanceof CellError) return rates;
  if (!Number.isFinite(base)) return webSheetFormulaError(ErrorType.VALUE);
  return rates.reduce((value, rate) => value * (1 + rate), base);
}



function webSheetFormulaXnpv(rate, values, dates) {
  const r = Number(rate);
  if (!Number.isFinite(r) || r <= -1) return webSheetFormulaError(ErrorType.NUM);
  const cashFlows = webSheetFormulaFinancialValues(values);
  if (cashFlows instanceof CellError) return cashFlows;
  const dateValues = webSheetFormulaFlattenValues(dates).map((value) => webSheetFormulaDateSerial(value));
  const dateError = dateValues.find((value) => value instanceof CellError);
  if (dateError) return dateError;
  if (!cashFlows.length || cashFlows.length !== dateValues.length) return webSheetFormulaError(ErrorType.NUM);
  const firstDate = dateValues[0];
  if (!Number.isFinite(firstDate)) return webSheetFormulaError(ErrorType.VALUE);
  return cashFlows.reduce((total, value, index) => {
    const days = dateValues[index] - firstDate;
    return total + value / ((1 + r) ** (days / 365));
  }, 0);
}



function webSheetFormulaXirr(values, dates, guess = 0.1) {
  const cashFlows = webSheetFormulaFinancialValues(values);
  if (cashFlows instanceof CellError) return cashFlows;
  const dateValues = webSheetFormulaFlattenValues(dates).map((value) => webSheetFormulaDateSerial(value));
  const dateError = dateValues.find((value) => value instanceof CellError);
  if (dateError) return dateError;
  if (
    cashFlows.length < 2 ||
    cashFlows.length !== dateValues.length ||
    !cashFlows.some((value) => value > 0) ||
    !cashFlows.some((value) => value < 0)
  ) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const firstDate = dateValues[0];
  if (!Number.isFinite(firstDate)) return webSheetFormulaError(ErrorType.VALUE);
  const valueAt = (rate) => cashFlows.reduce((sum, value, index) =>
    sum + value / ((1 + rate) ** ((dateValues[index] - firstDate) / 365)),
  0);
  const derivativeAt = (rate) => cashFlows.reduce((sum, value, index) => {
    const power = (dateValues[index] - firstDate) / 365;
    return sum - (power * value) / ((1 + rate) ** (power + 1));
  }, 0);
  let rate = Number.isFinite(Number(guess)) ? Number(guess) : 0.1;
  for (let index = 0; index < 100; index += 1) {
    const value = valueAt(rate);
    const derivative = derivativeAt(rate);
    if (!Number.isFinite(value) || !Number.isFinite(derivative) || Math.abs(derivative) < 1e-12) break;
    const next = rate - value / derivative;
    if (!Number.isFinite(next) || next <= -0.999999999) break;
    if (Math.abs(next - rate) < 1e-10) return next;
    rate = next;
  }
  let low = -0.999999;
  let high = Math.max(1, rate || 1);
  let fLow = valueAt(low);
  let fHigh = valueAt(high);
  for (let index = 0; index < 80 && Math.sign(fLow) === Math.sign(fHigh); index += 1) {
    high *= 2;
    fHigh = valueAt(high);
  }
  if (!Number.isFinite(fLow) || !Number.isFinite(fHigh) || Math.sign(fLow) === Math.sign(fHigh)) return webSheetFormulaError(ErrorType.NUM);
  for (let index = 0; index < 120; index += 1) {
    const mid = (low + high) / 2;
    const fMid = valueAt(mid);
    if (!Number.isFinite(fMid)) return webSheetFormulaError(ErrorType.NUM);
    if (Math.abs(fMid) < 1e-8 || Math.abs(high - low) < 1e-10) return mid;
    if (Math.sign(fLow) === Math.sign(fMid)) {
      low = mid;
      fLow = fMid;
    } else {
      high = mid;
    }
  }
  return webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaSecurityYearFraction(settlement, maturity, basis = 0) {
  const start = webSheetFormulaDateSerial(settlement);
  const end = webSheetFormulaDateSerial(maturity);
  const dayBasis = Math.trunc(Number(basis == null || basis === "" ? 0 : basis));
  if (start instanceof CellError) return start;
  if (end instanceof CellError) return end;
  if (end <= start || ![0, 1, 2, 3, 4].includes(dayBasis)) return webSheetFormulaError(ErrorType.NUM);
  return webSheetFormulaYearFrac(start, end, dayBasis);
}



function webSheetFormulaDisc(settlement, maturity, pr, redemption, basis = 0) {
  const fraction = webSheetFormulaSecurityYearFraction(settlement, maturity, basis);
  const price = Number(pr);
  const redeem = Number(redemption);
  if (fraction instanceof CellError) return fraction;
  if (![price, redeem].every(Number.isFinite) || price <= 0 || redeem <= 0 || fraction <= 0) return webSheetFormulaError(ErrorType.NUM);
  return (redeem - price) / (redeem * fraction);
}



function webSheetFormulaPriceDisc(settlement, maturity, discount, redemption, basis = 0) {
  const fraction = webSheetFormulaSecurityYearFraction(settlement, maturity, basis);
  const rate = Number(discount);
  const redeem = Number(redemption);
  if (fraction instanceof CellError) return fraction;
  if (![rate, redeem].every(Number.isFinite) || rate <= 0 || redeem <= 0 || fraction <= 0) return webSheetFormulaError(ErrorType.NUM);
  const price = redeem * (1 - rate * fraction);
  return price > 0 ? price : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaYieldDisc(settlement, maturity, pr, redemption, basis = 0) {
  const fraction = webSheetFormulaSecurityYearFraction(settlement, maturity, basis);
  const price = Number(pr);
  const redeem = Number(redemption);
  if (fraction instanceof CellError) return fraction;
  if (![price, redeem].every(Number.isFinite) || price <= 0 || redeem <= 0 || fraction <= 0) return webSheetFormulaError(ErrorType.NUM);
  return (redeem - price) / (price * fraction);
}



function webSheetFormulaIntrate(settlement, maturity, investment, redemption, basis = 0) {
  const fraction = webSheetFormulaSecurityYearFraction(settlement, maturity, basis);
  const invest = Number(investment);
  const redeem = Number(redemption);
  if (fraction instanceof CellError) return fraction;
  if (![invest, redeem].every(Number.isFinite) || invest <= 0 || redeem <= 0 || fraction <= 0) return webSheetFormulaError(ErrorType.NUM);
  return (redeem - invest) / (invest * fraction);
}



function webSheetFormulaReceived(settlement, maturity, investment, discount, basis = 0) {
  const fraction = webSheetFormulaSecurityYearFraction(settlement, maturity, basis);
  const invest = Number(investment);
  const rate = Number(discount);
  if (fraction instanceof CellError) return fraction;
  if (![invest, rate].every(Number.isFinite) || invest <= 0 || rate <= 0 || fraction <= 0) return webSheetFormulaError(ErrorType.NUM);
  const denominator = 1 - rate * fraction;
  return denominator > 0 ? invest / denominator : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaSecurityDateSet(issue, settlement, maturity, basis = 0) {
  const issued = webSheetFormulaDateSerial(issue);
  const settle = webSheetFormulaDateSerial(settlement);
  const mature = webSheetFormulaDateSerial(maturity);
  const dayBasis = webSheetFormulaSecurityBasis(basis);
  if (issued instanceof CellError) return issued;
  if (settle instanceof CellError) return settle;
  if (mature instanceof CellError) return mature;
  if (dayBasis == null || issued >= settle || settle >= mature) return webSheetFormulaError(ErrorType.NUM);
  return { issue: issued, settlement: settle, maturity: mature, basis: dayBasis };
}



function webSheetFormulaAccrint(issue, firstInterest, settlement, rate, par = 1000, frequency = 1, basis = 0, calcMethod = true) {
  const issued = webSheetFormulaDateSerial(issue);
  const first = webSheetFormulaDateSerial(firstInterest);
  const settle = webSheetFormulaDateSerial(settlement);
  const couponRate = Number(rate);
  const principal = Number(par == null || par === "" ? 1000 : par);
  const freq = webSheetFormulaSecurityFrequency(frequency);
  const dayBasis = webSheetFormulaSecurityBasis(basis);
  if (issued instanceof CellError) return issued;
  if (first instanceof CellError) return first;
  if (settle instanceof CellError) return settle;
  if (!freq || dayBasis == null || issued >= first || issued >= settle || first <= issued || !Number.isFinite(couponRate) || !Number.isFinite(principal) || couponRate <= 0 || principal <= 0) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const useIssueDate = calcMethod === true || calcMethod === 1 || String(calcMethod).toUpperCase() === "TRUE";
  const accrualStart = useIssueDate || settle <= first ? issued : first;
  const fraction = webSheetFormulaYearFrac(accrualStart, settle, dayBasis);
  return fraction instanceof CellError ? fraction : principal * couponRate * fraction;
}



function webSheetFormulaAccrintM(issue, settlement, rate, par = 1000, basis = 0) {
  const issued = webSheetFormulaDateSerial(issue);
  const settle = webSheetFormulaDateSerial(settlement);
  const couponRate = Number(rate);
  const principal = Number(par == null || par === "" ? 1000 : par);
  const dayBasis = webSheetFormulaSecurityBasis(basis);
  if (issued instanceof CellError) return issued;
  if (settle instanceof CellError) return settle;
  if (dayBasis == null || issued >= settle || !Number.isFinite(couponRate) || !Number.isFinite(principal) || couponRate <= 0 || principal <= 0) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const fraction = webSheetFormulaYearFrac(issued, settle, dayBasis);
  return fraction instanceof CellError ? fraction : principal * couponRate * fraction;
}



function webSheetFormulaPriceMat(settlement, maturity, issue, rate, yld, basis = 0) {
  const dates = webSheetFormulaSecurityDateSet(issue, settlement, maturity, basis);
  const couponRate = Number(rate);
  const yieldRate = Number(yld);
  if (dates instanceof CellError) return dates;
  if (![couponRate, yieldRate].every(Number.isFinite) || couponRate < 0 || yieldRate < 0) return webSheetFormulaError(ErrorType.NUM);
  const issueToMaturity = webSheetFormulaYearFrac(dates.issue, dates.maturity, dates.basis);
  const issueToSettlement = webSheetFormulaYearFrac(dates.issue, dates.settlement, dates.basis);
  const settlementToMaturity = webSheetFormulaYearFrac(dates.settlement, dates.maturity, dates.basis);
  if ([issueToMaturity, issueToSettlement, settlementToMaturity].some((value) => value instanceof CellError)) {
    return issueToMaturity instanceof CellError ? issueToMaturity : issueToSettlement instanceof CellError ? issueToSettlement : settlementToMaturity;
  }
  const denominator = 1 + yieldRate * settlementToMaturity;
  if (denominator <= 0) return webSheetFormulaError(ErrorType.NUM);
  return (100 + 100 * couponRate * issueToMaturity) / denominator - 100 * couponRate * issueToSettlement;
}



function webSheetFormulaYieldMat(settlement, maturity, issue, rate, pr, basis = 0) {
  const dates = webSheetFormulaSecurityDateSet(issue, settlement, maturity, basis);
  const couponRate = Number(rate);
  const price = Number(pr);
  if (dates instanceof CellError) return dates;
  if (![couponRate, price].every(Number.isFinite) || couponRate < 0 || price <= 0) return webSheetFormulaError(ErrorType.NUM);
  const issueToMaturity = webSheetFormulaYearFrac(dates.issue, dates.maturity, dates.basis);
  const issueToSettlement = webSheetFormulaYearFrac(dates.issue, dates.settlement, dates.basis);
  const settlementToMaturity = webSheetFormulaYearFrac(dates.settlement, dates.maturity, dates.basis);
  if ([issueToMaturity, issueToSettlement, settlementToMaturity].some((value) => value instanceof CellError)) {
    return issueToMaturity instanceof CellError ? issueToMaturity : issueToSettlement instanceof CellError ? issueToSettlement : settlementToMaturity;
  }
  const adjustedPrice = price + 100 * couponRate * issueToSettlement;
  return adjustedPrice > 0 && settlementToMaturity > 0
    ? (((100 + 100 * couponRate * issueToMaturity) / adjustedPrice) - 1) / settlementToMaturity
    : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaSecurityFrequency(frequency) {
  const value = Math.trunc(Number(frequency));
  return [1, 2, 4].includes(value) ? value : null;
}



function webSheetFormulaSecurityBasis(basis = 0) {
  const value = Math.trunc(Number(basis == null || basis === "" ? 0 : basis));
  return [0, 1, 2, 3, 4].includes(value) ? value : null;
}



function webSheetFormulaCouponPartialDays(start, end, basis) {
  if (basis === 0) return webSheetFormulaDays360(start, end, false);
  if (basis === 4) return webSheetFormulaDays360(start, end, true);
  return Math.trunc(Number(end)) - Math.trunc(Number(start));
}



function webSheetFormulaCouponPeriodDays(previous, next, basis, frequency) {
  if (basis === 0 || basis === 2 || basis === 4) return 360 / frequency;
  if (basis === 3) return 365 / frequency;
  return Math.trunc(Number(next)) - Math.trunc(Number(previous));
}



function webSheetFormulaCouponInfo(settlement, maturity, frequency, basis = 0) {
  const settle = webSheetFormulaDateSerial(settlement);
  const mature = webSheetFormulaDateSerial(maturity);
  const freq = webSheetFormulaSecurityFrequency(frequency);
  const dayBasis = webSheetFormulaSecurityBasis(basis);
  if (settle instanceof CellError) return settle;
  if (mature instanceof CellError) return mature;
  if (!freq || dayBasis == null || settle >= mature) return webSheetFormulaError(ErrorType.NUM);
  const months = 12 / freq;
  let next = mature;
  let previous = webSheetFormulaAddMonths(next, -months);
  let couponCount = 1;
  while (!(previous instanceof CellError) && previous > settle) {
    next = previous;
    previous = webSheetFormulaAddMonths(previous, -months);
    couponCount += 1;
  }
  if (previous instanceof CellError) return previous;
  const days = webSheetFormulaCouponPeriodDays(previous, next, dayBasis, freq);
  const daysBefore = webSheetFormulaCouponPartialDays(previous, settle, dayBasis);
  const daysAfter = webSheetFormulaCouponPartialDays(settle, next, dayBasis);
  if (![days, daysBefore, daysAfter].every(Number.isFinite) || days <= 0 || daysBefore < 0 || daysAfter < 0) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  return { settlement: settle, maturity: mature, frequency: freq, basis: dayBasis, previous, next, couponCount, days, daysBefore, daysAfter };
}



function webSheetFormulaCoupDayBs(settlement, maturity, frequency, basis = 0) {
  const info = webSheetFormulaCouponInfo(settlement, maturity, frequency, basis);
  return info instanceof CellError ? info : info.daysBefore;
}



function webSheetFormulaCoupDays(settlement, maturity, frequency, basis = 0) {
  const info = webSheetFormulaCouponInfo(settlement, maturity, frequency, basis);
  return info instanceof CellError ? info : info.days;
}



function webSheetFormulaCoupDaysNc(settlement, maturity, frequency, basis = 0) {
  const info = webSheetFormulaCouponInfo(settlement, maturity, frequency, basis);
  return info instanceof CellError ? info : info.daysAfter;
}



function webSheetFormulaCoupNcd(settlement, maturity, frequency, basis = 0) {
  const info = webSheetFormulaCouponInfo(settlement, maturity, frequency, basis);
  return info instanceof CellError ? info : info.next;
}



function webSheetFormulaCoupPcd(settlement, maturity, frequency, basis = 0) {
  const info = webSheetFormulaCouponInfo(settlement, maturity, frequency, basis);
  return info instanceof CellError ? info : info.previous;
}



function webSheetFormulaCoupNum(settlement, maturity, frequency, basis = 0) {
  const info = webSheetFormulaCouponInfo(settlement, maturity, frequency, basis);
  return info instanceof CellError ? info : info.couponCount;
}



function webSheetFormulaSecurityPrice(settlement, maturity, rate, yld, redemption, frequency, basis = 0) {
  const info = webSheetFormulaCouponInfo(settlement, maturity, frequency, basis);
  const couponRate = Number(rate);
  const yieldRate = Number(yld);
  const redeem = Number(redemption);
  if (info instanceof CellError) return info;
  if (![couponRate, yieldRate, redeem].every(Number.isFinite) || couponRate < 0 || yieldRate < 0 || redeem <= 0) return webSheetFormulaError(ErrorType.NUM);
  const coupon = 100 * couponRate / info.frequency;
  const yieldFactor = 1 + yieldRate / info.frequency;
  const offset = info.daysAfter / info.days;
  const accrued = coupon * info.daysBefore / info.days;
  let dirtyPrice = redeem / (yieldFactor ** (info.couponCount - 1 + offset));
  for (let couponIndex = 0; couponIndex < info.couponCount; couponIndex += 1) {
    dirtyPrice += coupon / (yieldFactor ** (couponIndex + offset));
  }
  return dirtyPrice - accrued;
}



function webSheetFormulaSecurityYield(settlement, maturity, rate, pr, redemption, frequency, basis = 0) {
  const price = Number(pr);
  const couponRate = Number(rate);
  const redeem = Number(redemption);
  const freq = webSheetFormulaSecurityFrequency(frequency);
  if (![price, couponRate, redeem].every(Number.isFinite) || price <= 0 || couponRate < 0 || redeem <= 0 || !freq) return webSheetFormulaError(ErrorType.NUM);
  const valueAt = (candidate) => {
    const value = webSheetFormulaSecurityPrice(settlement, maturity, couponRate, candidate, redeem, freq, basis);
    return value instanceof CellError ? value : value - price;
  };
  let low = 0;
  let high = Math.max(1, couponRate || 0.1);
  let fLow = valueAt(low);
  let fHigh = valueAt(high);
  if (fLow instanceof CellError) return fLow;
  if (fHigh instanceof CellError) return fHigh;
  for (let index = 0; index < 80 && Math.sign(fLow) === Math.sign(fHigh); index += 1) {
    high *= 2;
    fHigh = valueAt(high);
    if (fHigh instanceof CellError) return fHigh;
  }
  if (!Number.isFinite(fLow) || !Number.isFinite(fHigh) || Math.sign(fLow) === Math.sign(fHigh)) return webSheetFormulaError(ErrorType.NUM);
  for (let index = 0; index < 120; index += 1) {
    const mid = (low + high) / 2;
    const fMid = valueAt(mid);
    if (fMid instanceof CellError) return fMid;
    if (Math.abs(fMid) < 1e-8 || Math.abs(high - low) < 1e-10) return mid;
    if (Math.sign(fLow) === Math.sign(fMid)) {
      low = mid;
      fLow = fMid;
    } else {
      high = mid;
    }
  }
  return webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaSecurityDuration(settlement, maturity, coupon, yld, frequency, basis = 0, modified = false) {
  const info = webSheetFormulaCouponInfo(settlement, maturity, frequency, basis);
  const couponRate = Number(coupon);
  const yieldRate = Number(yld);
  if (info instanceof CellError) return info;
  if (![couponRate, yieldRate].every(Number.isFinite) || couponRate < 0 || yieldRate < 0) return webSheetFormulaError(ErrorType.NUM);
  const couponPayment = 100 * couponRate / info.frequency;
  const yieldFactor = 1 + yieldRate / info.frequency;
  const offset = info.daysAfter / info.days;
  let presentValueTotal = 0;
  let weightedTimeTotal = 0;
  for (let couponIndex = 0; couponIndex < info.couponCount; couponIndex += 1) {
    const period = couponIndex + offset;
    const years = period / info.frequency;
    const cashFlow = couponPayment + (couponIndex === info.couponCount - 1 ? 100 : 0);
    const presentValue = cashFlow / (yieldFactor ** period);
    presentValueTotal += presentValue;
    weightedTimeTotal += years * presentValue;
  }
  if (presentValueTotal <= 0) return webSheetFormulaError(ErrorType.NUM);
  const duration = weightedTimeTotal / presentValueTotal;
  return modified ? duration / yieldFactor : duration;
}



function webSheetFormulaIrr(values, guess = 0.1) {
  const numbers = webSheetFormulaFinancialValues(values);
  if (numbers instanceof CellError) return numbers;
  if (numbers.length < 2 || !numbers.some((value) => value > 0) || !numbers.some((value) => value < 0)) return webSheetFormulaError(ErrorType.NUM);
  const valueAt = (rate) => numbers.reduce((sum, value, index) => sum + value / ((1 + rate) ** index), 0);
  let rate = Number.isFinite(Number(guess)) ? Number(guess) : 0.1;
  for (let index = 0; index < 100; index += 1) {
    let f = 0;
    let derivative = 0;
    for (let period = 0; period < numbers.length; period += 1) {
      const factor = (1 + rate) ** period;
      f += numbers[period] / factor;
      if (period) derivative -= period * numbers[period] / ((1 + rate) ** (period + 1));
    }
    const next = rate - f / derivative;
    if (!Number.isFinite(next) || next <= -0.999999999) break;
    if (Math.abs(next - rate) < 1e-10) return next;
    rate = next;
  }
  let low = -0.999999;
  let high = Math.max(1, Number(guess) || 1);
  let fLow = valueAt(low);
  let fHigh = valueAt(high);
  for (let index = 0; index < 80 && Math.sign(fLow) === Math.sign(fHigh); index += 1) {
    high *= 2;
    fHigh = valueAt(high);
  }
  if (!Number.isFinite(fLow) || !Number.isFinite(fHigh) || Math.sign(fLow) === Math.sign(fHigh)) return webSheetFormulaError(ErrorType.NUM);
  for (let index = 0; index < 100; index += 1) {
    const mid = (low + high) / 2;
    const fMid = valueAt(mid);
    if (Math.abs(fMid) < 1e-8 || Math.abs(high - low) < 1e-10) return mid;
    if (Math.sign(fLow) === Math.sign(fMid)) {
      low = mid;
      fLow = fMid;
    } else {
      high = mid;
    }
  }
  return webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaMirr(values, financeRate, reinvestRate) {
  const numbers = webSheetFormulaFinancialValues(values);
  if (numbers instanceof CellError) return numbers;
  if (numbers.length < 2 || !numbers.some((value) => value > 0) || !numbers.some((value) => value < 0)) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const finance = Number(financeRate);
  const reinvest = Number(reinvestRate);
  if (!Number.isFinite(finance) || !Number.isFinite(reinvest) || finance <= -1 || reinvest <= -1) return webSheetFormulaError(ErrorType.NUM);
  const finalPeriod = numbers.length - 1;
  let negativePv = 0;
  let positiveFv = 0;
  numbers.forEach((value, index) => {
    if (value < 0) negativePv += value / ((1 + finance) ** index);
    else if (value > 0) positiveFv += value * ((1 + reinvest) ** (finalPeriod - index));
  });
  if (!negativePv || !positiveFv) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  return ((-positiveFv / negativePv) ** (1 / finalPeriod)) - 1;
}



function webSheetFormulaSln(cost, salvage, life) {
  const values = [Number(cost), Number(salvage), Number(life)];
  if (!values.every(Number.isFinite) || values[2] <= 0) return webSheetFormulaError(ErrorType.NUM);
  return (values[0] - values[1]) / values[2];
}



function webSheetFormulaSyd(cost, salvage, life, per) {
  const assetCost = Number(cost);
  const assetSalvage = Number(salvage);
  const assetLife = Number(life);
  const period = Number(per);
  if (![assetCost, assetSalvage, assetLife, period].every(Number.isFinite) || assetLife <= 0 || period <= 0 || period > assetLife) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  return ((assetCost - assetSalvage) * (assetLife - period + 1) * 2) / (assetLife * (assetLife + 1));
}



function webSheetFormulaDdb(cost, salvage, life, period, factor = 2) {
  const assetCost = Number(cost);
  const assetSalvage = Number(salvage);
  const assetLife = Number(life);
  const targetPeriod = Math.trunc(Number(period));
  const depreciationFactor = Number(factor);
  if (
    ![assetCost, assetSalvage, assetLife, depreciationFactor].every(Number.isFinite) ||
    assetCost < 0 ||
    assetSalvage < 0 ||
    assetLife <= 0 ||
    targetPeriod < 1 ||
    targetPeriod > assetLife ||
    depreciationFactor <= 0
  ) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  let bookValue = assetCost;
  let depreciation = 0;
  for (let current = 1; current <= targetPeriod; current += 1) {
    depreciation = Math.min(bookValue * depreciationFactor / assetLife, Math.max(0, bookValue - assetSalvage));
    bookValue -= depreciation;
  }
  return depreciation;
}



function webSheetFormulaDb(cost, salvage, life, period, month = 12) {
  const assetCost = Number(cost);
  const assetSalvage = Number(salvage);
  const assetLife = Number(life);
  const targetPeriod = Math.trunc(Number(period));
  const firstYearMonths = Math.trunc(Number(month == null || month === "" ? 12 : month));
  if (
    ![assetCost, assetSalvage, assetLife].every(Number.isFinite) ||
    assetCost <= 0 ||
    assetSalvage < 0 ||
    assetLife <= 0 ||
    targetPeriod < 1 ||
    targetPeriod > assetLife + 1 ||
    firstYearMonths < 1 ||
    firstYearMonths > 12
  ) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const rate = Math.round((1 - ((assetSalvage / assetCost) ** (1 / assetLife))) * 1000) / 1000;
  let accumulated = 0;
  let depreciation = 0;
  for (let current = 1; current <= targetPeriod; current += 1) {
    const bookValue = assetCost - accumulated;
    if (current === 1) depreciation = bookValue * rate * firstYearMonths / 12;
    else if (current === Math.floor(assetLife) + 1) depreciation = bookValue * rate * (12 - firstYearMonths) / 12;
    else depreciation = bookValue * rate;
    const remainingDepreciable = Math.max(0, assetCost - assetSalvage - accumulated);
    depreciation = Math.min(depreciation, remainingDepreciable);
    accumulated += depreciation;
  }
  return depreciation;
}



function webSheetFormulaVdb(cost, salvage, life, startPeriod, endPeriod, factor = 2, noSwitch = false) {
  const assetCost = Number(cost);
  const assetSalvage = Number(salvage);
  const assetLife = Number(life);
  const start = Number(startPeriod);
  const end = Number(endPeriod);
  const depreciationFactor = Number(factor == null || factor === "" ? 2 : factor);
  if (
    ![assetCost, assetSalvage, assetLife, start, end, depreciationFactor].every(Number.isFinite) ||
    assetCost < 0 ||
    assetSalvage < 0 ||
    assetLife <= 0 ||
    start < 0 ||
    end < start ||
    depreciationFactor <= 0
  ) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const useDecliningOnly = noSwitch === true || noSwitch === 1 || String(noSwitch).toUpperCase() === "TRUE";
  let bookValue = assetCost;
  let total = 0;
  const finalPeriod = Math.ceil(end);
  for (let period = 1; period <= finalPeriod; period += 1) {
    const remainingPeriods = Math.max(assetLife - period + 1, 1);
    const declining = Math.min(bookValue * depreciationFactor / assetLife, Math.max(0, bookValue - assetSalvage));
    const straightLine = Math.max(0, (bookValue - assetSalvage) / remainingPeriods);
    const depreciation = useDecliningOnly ? declining : Math.max(declining, straightLine);
    const overlap = Math.max(0, Math.min(end, period) - Math.max(start, period - 1));
    if (overlap > 0) total += depreciation * overlap;
    bookValue = Math.max(assetSalvage, bookValue - depreciation);
  }
  return total;
}



function webSheetFormulaAmorFirstFraction(datePurchased, firstPeriod, basis = 0) {
  const purchased = webSheetFormulaDateSerial(datePurchased);
  const first = webSheetFormulaDateSerial(firstPeriod);
  const dayBasis = webSheetFormulaSecurityBasis(basis);
  if (purchased instanceof CellError) return purchased;
  if (first instanceof CellError) return first;
  if (dayBasis == null || first <= purchased) return webSheetFormulaError(ErrorType.NUM);
  return webSheetFormulaYearFrac(purchased, first, dayBasis);
}



function webSheetFormulaAmorLinc(cost, datePurchased, firstPeriod, salvage, period, rate, basis = 0) {
  const assetCost = Number(cost);
  const assetSalvage = Number(salvage);
  const targetPeriod = Math.trunc(Number(period));
  const depreciationRate = Number(rate);
  const firstFraction = webSheetFormulaAmorFirstFraction(datePurchased, firstPeriod, basis);
  if (firstFraction instanceof CellError) return firstFraction;
  if (
    ![assetCost, assetSalvage, targetPeriod, depreciationRate].every(Number.isFinite) ||
    assetCost <= 0 ||
    assetSalvage < 0 ||
    assetSalvage >= assetCost ||
    targetPeriod < 0 ||
    depreciationRate <= 0
  ) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  let accumulated = 0;
  let depreciation = 0;
  for (let current = 0; current <= targetPeriod; current += 1) {
    const baseDepreciation = assetCost * depreciationRate * (current === 0 ? firstFraction : 1);
    const remaining = Math.max(0, assetCost - assetSalvage - accumulated);
    depreciation = Math.min(baseDepreciation, remaining);
    accumulated += depreciation;
  }
  return depreciation;
}



function webSheetFormulaAmorDegrcCoefficient(life) {
  if (life < 3) return 1;
  if (life < 5) return 1.5;
  if (life <= 6) return 2;
  return 2.5;
}



function webSheetFormulaAmorDegrc(cost, datePurchased, firstPeriod, salvage, period, rate, basis = 0) {
  const assetCost = Number(cost);
  const assetSalvage = Number(salvage);
  const targetPeriod = Math.trunc(Number(period));
  const depreciationRate = Number(rate);
  const firstFraction = webSheetFormulaAmorFirstFraction(datePurchased, firstPeriod, basis);
  if (firstFraction instanceof CellError) return firstFraction;
  if (
    ![assetCost, assetSalvage, targetPeriod, depreciationRate].every(Number.isFinite) ||
    assetCost <= 0 ||
    assetSalvage < 0 ||
    assetSalvage >= assetCost ||
    targetPeriod < 0 ||
    depreciationRate <= 0
  ) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const coefficient = webSheetFormulaAmorDegrcCoefficient(1 / depreciationRate);
  const adjustedRate = depreciationRate * coefficient;
  let bookValue = assetCost;
  let depreciation = 0;
  for (let current = 0; current <= targetPeriod; current += 1) {
    const fraction = current === 0 ? firstFraction : 1;
    const remaining = Math.max(0, bookValue - assetSalvage);
    depreciation = Math.min(bookValue * adjustedRate * fraction, remaining);
    bookValue -= depreciation;
  }
  return depreciation;
}



function webSheetFormulaDollarFractionDenominator(fraction) {
  const value = Math.trunc(Number(fraction));
  return Number.isFinite(value) && value >= 1 ? value : null;
}



function webSheetFormulaDollarDe(fractionalDollar, fraction) {
  const denominator = webSheetFormulaDollarFractionDenominator(fraction);
  const value = Number(fractionalDollar);
  if (!Number.isFinite(value) || denominator == null) return webSheetFormulaError(ErrorType.NUM);
  const sign = value < 0 ? -1 : 1;
  const absolute = Math.abs(value);
  const whole = Math.trunc(absolute);
  const decimalText = String(absolute).split(".")[1] || "";
  const numerator = decimalText ? Number(decimalText) : 0;
  return sign * (whole + numerator / denominator);
}



function webSheetFormulaDollarFr(decimalDollar, fraction) {
  const denominator = webSheetFormulaDollarFractionDenominator(fraction);
  const value = Number(decimalDollar);
  if (!Number.isFinite(value) || denominator == null) return webSheetFormulaError(ErrorType.NUM);
  const sign = value < 0 ? -1 : 1;
  const absolute = Math.abs(value);
  const whole = Math.trunc(absolute);
  const numerator = Math.round((absolute - whole) * denominator);
  const scale = 10 ** String(denominator).length;
  return sign * (whole + numerator / scale);
}



function webSheetFormulaRri(nper, pv, fv) {
  const periods = Number(nper);
  const present = Number(pv);
  const future = Number(fv);
  if (![periods, present, future].every(Number.isFinite) || periods <= 0 || present <= 0 || future <= 0) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  return (future / present) ** (1 / periods) - 1;
}



function webSheetFormulaPduration(rate, pv, fv) {
  const r = Number(rate);
  const present = Number(pv);
  const future = Number(fv);
  if (![r, present, future].every(Number.isFinite) || r <= 0 || present <= 0 || future <= 0) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  return Math.log(future / present) / Math.log(1 + r);
}



function webSheetFormulaIspmt(rate, per, nper, pv) {
  const r = Number(rate);
  const period = Number(per);
  const periods = Number(nper);
  const present = Number(pv);
  if (![r, period, periods, present].every(Number.isFinite) || periods <= 0 || period < 1 || period > periods) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  return r * (period - periods) * present / periods;
}



function webSheetFormulaTbillDays(settlement, maturity) {
  const start = webSheetFormulaDateSerial(settlement);
  const end = webSheetFormulaDateSerial(maturity);
  if (start instanceof CellError) return start;
  if (end instanceof CellError) return end;
  const days = end - start;
  return days > 0 && days <= 365 ? days : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaTbillEq(settlement, maturity, discount) {
  const days = webSheetFormulaTbillDays(settlement, maturity);
  const rate = Number(discount);
  if (days instanceof CellError) return days;
  if (!Number.isFinite(rate) || rate <= 0) return webSheetFormulaError(ErrorType.NUM);
  return (365 * rate) / (360 - rate * days);
}



function webSheetFormulaTbillPrice(settlement, maturity, discount) {
  const days = webSheetFormulaTbillDays(settlement, maturity);
  const rate = Number(discount);
  if (days instanceof CellError) return days;
  if (!Number.isFinite(rate) || rate <= 0) return webSheetFormulaError(ErrorType.NUM);
  const price = 100 * (1 - rate * days / 360);
  return price > 0 ? price : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaTbillYield(settlement, maturity, price) {
  const days = webSheetFormulaTbillDays(settlement, maturity);
  const value = Number(price);
  if (days instanceof CellError) return days;
  if (!Number.isFinite(value) || value <= 0) return webSheetFormulaError(ErrorType.NUM);
  return ((100 - value) / value) * (360 / days);
}



function webSheetFormulaIpmt(rate, per, nper, pv, fv = 0, type = 0) {
  const period = Math.trunc(Number(per));
  const periods = Number(nper);
  const paymentType = webSheetFormulaFinancialType(type);
  if (!Number.isFinite(periods) || period < 1 || period > periods || paymentType == null) return webSheetFormulaError(ErrorType.NUM);
  const payment = webSheetFormulaPayment(rate, periods, pv, fv, paymentType);
  if (payment instanceof CellError) return payment;
  const r = Number(rate);
  if (r === 0) return 0;
  if (paymentType === 1 && period === 1) return 0;
  const balancePeriod = paymentType === 1 ? period - 2 : period - 1;
  const balance = webSheetFormulaFutureValue(r, balancePeriod, payment, pv, paymentType);
  return balance instanceof CellError ? balance : balance * r;
}



function webSheetFormulaPpmt(rate, per, nper, pv, fv = 0, type = 0) {
  const payment = webSheetFormulaPayment(rate, nper, pv, fv, type);
  if (payment instanceof CellError) return payment;
  const interest = webSheetFormulaIpmt(rate, per, nper, pv, fv, type);
  return interest instanceof CellError ? interest : payment - interest;
}



function webSheetFormulaRoman(number) {
  let value = Math.trunc(Number(number));
  if (!Number.isFinite(value) || value <= 0 || value >= 4000) return webSheetFormulaError(ErrorType.VALUE);
  let output = "";
  for (const [amount, symbol] of WEBSHEET_ROMAN_VALUES) {
    while (value >= amount) {
      output += symbol;
      value -= amount;
    }
  }
  return output;
}



function webSheetFormulaArabic(text) {
  const source = String(text ?? "").trim().toUpperCase();
  if (!source) return webSheetFormulaError(ErrorType.VALUE);
  const values = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  let total = 0;
  let previous = 0;
  for (let index = source.length - 1; index >= 0; index -= 1) {
    const value = values[source[index]];
    if (!value) return webSheetFormulaError(ErrorType.VALUE);
    total += value < previous ? -value : value;
    previous = Math.max(previous, value);
  }
  return webSheetFormulaRoman(total) instanceof CellError ? webSheetFormulaError(ErrorType.VALUE) : total;
}



function webSheetFormulaBase(number, radix, minLength = 0) {
  const value = Math.trunc(Number(number));
  const base = Math.trunc(Number(radix));
  const width = Math.trunc(Number(minLength || 0));
  if (!Number.isFinite(value) || !Number.isFinite(base) || !Number.isFinite(width) || value < 0 || base < 2 || base > 36 || width < 0) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const text = value.toString(base).toUpperCase();
  return text.padStart(width, "0");
}



function webSheetFormulaDecimal(text, radix) {
  const base = Math.trunc(Number(radix));
  const source = String(text ?? "").trim().toUpperCase();
  if (!source || !Number.isFinite(base) || base < 2 || base > 36) return webSheetFormulaError(ErrorType.NUM);
  const valid = /^[0-9A-Z]+$/u.test(source) && [...source].every((char) => parseInt(char, 36) < base);
  if (!valid) return webSheetFormulaError(ErrorType.NUM);
  return parseInt(source, base);
}



function webSheetFormulaPlacesValue(places) {
  if (places === WEBSHEET_FORMULA_NA_SENTINEL) return null;
  const value = Math.trunc(Number(webSheetFormulaRawValue(places)));
  return Number.isFinite(value) && value >= 0 && value <= 10 ? value : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaSignedBaseRange(base) {
  const modulus = base ** 10;
  return { modulus, min: -modulus / 2, max: modulus / 2 - 1 };
}



function webSheetFormulaParseSignedBase(value, base) {
  const source = String(webSheetFormulaRawValue(value) ?? "").trim().toUpperCase();
  if (!source || source.length > 10) return webSheetFormulaError(ErrorType.NUM);
  const digits = base === 2 ? /^[01]+$/u : base === 8 ? /^[0-7]+$/u : /^[0-9A-F]+$/u;
  if (!digits.test(source)) return webSheetFormulaError(ErrorType.NUM);
  const parsed = parseInt(source, base);
  if (!Number.isFinite(parsed)) return webSheetFormulaError(ErrorType.NUM);
  const { modulus } = webSheetFormulaSignedBaseRange(base);
  const negativeLead = base === 2 ? "1" : base === 8 ? "4" : "8";
  if (source.length === 10 && parseInt(source[0], base) >= parseInt(negativeLead, base)) return parsed - modulus;
  return parsed;
}



function webSheetFormulaDecimalToSignedBase(number, base, places = WEBSHEET_FORMULA_NA_SENTINEL) {
  const value = Math.trunc(Number(webSheetFormulaRawValue(number)));
  const { modulus, min, max } = webSheetFormulaSignedBaseRange(base);
  if (!Number.isFinite(value) || value < min || value > max) return webSheetFormulaError(ErrorType.NUM);
  const width = webSheetFormulaPlacesValue(places);
  if (width instanceof CellError) return width;
  if (value < 0) return Math.trunc(modulus + value).toString(base).toUpperCase();
  const text = value.toString(base).toUpperCase();
  if (width == null) return text;
  return text.length <= width ? text.padStart(width, "0") : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaConvertBase(value, fromBase, toBase, places = WEBSHEET_FORMULA_NA_SENTINEL) {
  const decimal = webSheetFormulaParseSignedBase(value, fromBase);
  if (decimal instanceof CellError) return decimal;
  if (toBase === 10) return decimal;
  const { min, max } = webSheetFormulaSignedBaseRange(toBase);
  if (decimal < min || decimal > max) return webSheetFormulaError(ErrorType.NUM);
  return webSheetFormulaDecimalToSignedBase(decimal, toBase, places);
}



function webSheetFormulaBitNumber(value) {
  const number = Math.trunc(Number(webSheetFormulaRawValue(value)));
  const limit = 2 ** 48;
  return Number.isFinite(number) && number >= 0 && number < limit ? number : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaBitwise(left, right, operation) {
  const first = webSheetFormulaBitNumber(left);
  const second = webSheetFormulaBitNumber(right);
  if (first instanceof CellError) return first;
  if (second instanceof CellError) return second;
  const result = operation(BigInt(first), BigInt(second));
  return Number(result);
}



function webSheetFormulaBitShift(number, shiftAmount, direction) {
  const value = webSheetFormulaBitNumber(number);
  const shift = Math.trunc(Number(webSheetFormulaRawValue(shiftAmount)));
  if (value instanceof CellError) return value;
  if (!Number.isFinite(shift) || Math.abs(shift) > 53) return webSheetFormulaError(ErrorType.NUM);
  const normalized = direction * shift;
  const result = normalized >= 0 ? BigInt(value) << BigInt(normalized) : BigInt(value) >> BigInt(-normalized);
  return result < 2n ** 48n ? Number(result) : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaImabs(value) {
  const complex = dataAnalysisParseComplexValue(webSheetFormulaRawValue(value));
  return complex ? Math.sqrt(complex.real ** 2 + complex.imag ** 2) : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaComplex(value) {
  const complex = dataAnalysisParseComplexValue(webSheetFormulaRawValue(value));
  return complex || webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaComplexTextOrError(complex) {
  return complex instanceof CellError ? complex : dataAnalysisComplexText(complex.real, complex.imag);
}



function webSheetFormulaComplexText(real, imag, suffix = "i") {
  const r = dataAnalysisCleanSignedZero(Number(real));
  const i = dataAnalysisCleanSignedZero(Number(imag));
  const normalizedSuffix = String(suffix || "i").toLowerCase();
  if (!Number.isFinite(r) || !Number.isFinite(i) || !["i", "j"].includes(normalizedSuffix)) {
    return webSheetFormulaError(ErrorType.VALUE);
  }
  if (i === 0) return String(r);
  if (r === 0) return `${i}${normalizedSuffix}`;
  return `${r}${i < 0 ? "" : "+"}${i}${normalizedSuffix}`;
}



function webSheetFormulaComplexDivideValues(left, right) {
  const denominator = right.real ** 2 + right.imag ** 2;
  if (!denominator) return webSheetFormulaError(ErrorType.NUM);
  return {
    real: (left.real * right.real + left.imag * right.imag) / denominator,
    imag: (left.imag * right.real - left.real * right.imag) / denominator,
  };
}



function webSheetFormulaComplexUnary(value, transform) {
  const complex = webSheetFormulaComplex(value);
  if (complex instanceof CellError) return complex;
  const result = transform(complex);
  return result instanceof CellError ? result : webSheetFormulaComplexTextOrError(result);
}



function webSheetFormulaImsum(values) {
  let real = 0;
  let imag = 0;
  for (const value of values) {
    const complex = webSheetFormulaComplex(value);
    if (complex instanceof CellError) return complex;
    real += complex.real;
    imag += complex.imag;
  }
  return dataAnalysisComplexText(real, imag);
}



function webSheetFormulaImproduct(values) {
  let real = 1;
  let imag = 0;
  for (const value of values) {
    const complex = webSheetFormulaComplex(value);
    if (complex instanceof CellError) return complex;
    const nextReal = real * complex.real - imag * complex.imag;
    const nextImag = real * complex.imag + imag * complex.real;
    real = nextReal;
    imag = nextImag;
  }
  return dataAnalysisComplexText(real, imag);
}



function webSheetFormulaBesselOrder(value) {
  const order = Math.trunc(Number(webSheetFormulaRawValue(value)));
  return Number.isFinite(order) && order >= 0 ? order : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaBesselJ0(x) {
  const value = Math.abs(Number(x));
  if (!Number.isFinite(value)) return NaN;
  if (value < 8) {
    const y = value * value;
    const numerator =
      57568490574 +
      y * (-13362590354 +
      y * (651619640.7 +
      y * (-11214424.18 +
      y * (77392.33017 +
      y * -184.9052456))));
    const denominator =
      57568490411 +
      y * (1029532985 +
      y * (9494680.718 +
      y * (59272.64853 +
      y * (267.8532712 +
      y))));
    return numerator / denominator;
  }
  const z = 8 / value;
  const y = z * z;
  const xx = value - 0.785398164;
  const first =
    1 +
    y * (-0.001098628627 +
    y * (0.00002734510407 +
    y * (-0.000002073370639 +
    y * 0.0000002093887211)));
  const second =
    -0.01562499995 +
    y * (0.0001430488765 +
    y * (-0.000006911147651 +
    y * (0.0000007621095161 -
    y * 0.0000000934945152)));
  return Math.sqrt(0.636619772 / value) * (Math.cos(xx) * first - z * Math.sin(xx) * second);
}



function webSheetFormulaBesselJ1(x) {
  const number = Number(x);
  const value = Math.abs(number);
  if (!Number.isFinite(value)) return NaN;
  let result;
  if (value < 8) {
    const y = value * value;
    const numerator =
      value * (72362614232 +
      y * (-7895059235 +
      y * (242396853.1 +
      y * (-2972611.439 +
      y * (15704.48260 +
      y * -30.16036606)))));
    const denominator =
      144725228442 +
      y * (2300535178 +
      y * (18583304.74 +
      y * (99447.43394 +
      y * (376.9991397 +
      y))));
    result = numerator / denominator;
  } else {
    const z = 8 / value;
    const y = z * z;
    const xx = value - 2.356194491;
    const first =
      1 +
      y * (0.00183105 +
      y * (-0.00003516396496 +
      y * (0.000002457520174 -
      y * 0.000000240337019)));
    const second =
      0.04687499995 +
      y * (-0.0002002690873 +
      y * (0.000008449199096 +
      y * (-0.00000088228987 +
      y * 0.000000105787412)));
    result = Math.sqrt(0.636619772 / value) * (Math.cos(xx) * first - z * Math.sin(xx) * second);
  }
  return number < 0 ? -result : result;
}



function webSheetFormulaBesselY0(x) {
  const value = Number(x);
  if (!Number.isFinite(value) || value <= 0) return NaN;
  if (value < 8) {
    const y = value * value;
    const numerator =
      -2957821389 +
      y * (7062834065 +
      y * (-512359803.6 +
      y * (10879881.29 +
      y * (-86327.92757 +
      y * 228.4622733))));
    const denominator =
      40076544269 +
      y * (745249964.8 +
      y * (7189466.438 +
      y * (47447.26470 +
      y * (226.1030244 +
      y))));
    return numerator / denominator + 0.636619772 * webSheetFormulaBesselJ0(value) * Math.log(value);
  }
  const z = 8 / value;
  const y = z * z;
  const xx = value - 0.785398164;
  const first =
    1 +
    y * (-0.001098628627 +
    y * (0.00002734510407 +
    y * (-0.000002073370639 +
    y * 0.0000002093887211)));
  const second =
    -0.01562499995 +
    y * (0.0001430488765 +
    y * (-0.000006911147651 +
    y * (0.0000007621095161 -
    y * 0.0000000934945152)));
  return Math.sqrt(0.636619772 / value) * (Math.sin(xx) * first + z * Math.cos(xx) * second);
}



function webSheetFormulaBesselY1(x) {
  const value = Number(x);
  if (!Number.isFinite(value) || value <= 0) return NaN;
  if (value < 8) {
    const y = value * value;
    const numerator =
      value * (-0.4900604943e13 +
      y * (0.1274392739e13 +
      y * (-0.5153438139e11 +
      y * (0.7349264551e9 +
      y * (-0.4237922726e7 +
      y * 0.8511937935e4)))));
    const denominator =
      0.2499580570e14 +
      y * (0.4244419664e12 +
      y * (0.3733650367e10 +
      y * (0.2245904002e8 +
      y * (0.1020426050e6 +
      y * (0.3549632885e3 + y)))));
    return numerator / denominator + 0.636619772 * (webSheetFormulaBesselJ1(value) * Math.log(value) - 1 / value);
  }
  const z = 8 / value;
  const y = z * z;
  const xx = value - 2.356194491;
  const first =
    1 +
    y * (0.00183105 +
    y * (-0.00003516396496 +
    y * (0.000002457520174 -
    y * 0.000000240337019)));
  const second =
    0.04687499995 +
    y * (-0.0002002690873 +
    y * (0.000008449199096 +
    y * (-0.00000088228987 +
    y * 0.000000105787412)));
  return Math.sqrt(0.636619772 / value) * (Math.sin(xx) * first + z * Math.cos(xx) * second);
}



function webSheetFormulaBesselI0(x) {
  const value = Math.abs(Number(x));
  if (!Number.isFinite(value)) return NaN;
  if (value < 3.75) {
    const y = (value / 3.75) ** 2;
    return 1 + y * (3.5156229 + y * (3.0899424 + y * (1.2067492 + y * (0.2659732 + y * (0.0360768 + y * 0.0045813)))));
  }
  const y = 3.75 / value;
  return (Math.exp(value) / Math.sqrt(value)) *
    (0.39894228 + y * (0.01328592 + y * (0.00225319 + y * (-0.00157565 + y * (0.00916281 +
    y * (-0.02057706 + y * (0.02635537 + y * (-0.01647633 + y * 0.00392377))))))));
}



function webSheetFormulaBesselI1(x) {
  const number = Number(x);
  const value = Math.abs(number);
  if (!Number.isFinite(value)) return NaN;
  let result;
  if (value < 3.75) {
    const y = (value / 3.75) ** 2;
    result = value * (0.5 + y * (0.87890594 + y * (0.51498869 + y * (0.15084934 + y * (0.02658733 + y * (0.00301532 + y * 0.00032411))))));
  } else {
    const y = 3.75 / value;
    result = (Math.exp(value) / Math.sqrt(value)) *
      (0.39894228 + y * (-0.03988024 + y * (-0.00362018 + y * (0.00163801 + y * (-0.01031555 +
      y * (0.02282967 + y * (-0.02895312 + y * (0.01787654 - y * 0.00420059))))))));
  }
  return number < 0 ? -result : result;
}



function webSheetFormulaBesselK0(x) {
  const value = Number(x);
  if (!Number.isFinite(value) || value <= 0) return NaN;
  if (value <= 2) {
    const y = (value * value) / 4;
    return -Math.log(value / 2) * webSheetFormulaBesselI0(value) +
      (-0.57721566 + y * (0.42278420 + y * (0.23069756 + y * (0.03488590 +
      y * (0.00262698 + y * (0.00010750 + y * 0.00000740))))));
  }
  const y = 2 / value;
  return (Math.exp(-value) / Math.sqrt(value)) *
    (1.25331414 + y * (-0.07832358 + y * (0.02189568 + y * (-0.01062446 + y * (0.00587872 +
    y * (-0.00251540 + y * 0.00053208))))));
}



function webSheetFormulaBesselK1(x) {
  const value = Number(x);
  if (!Number.isFinite(value) || value <= 0) return NaN;
  if (value <= 2) {
    const y = (value * value) / 4;
    return Math.log(value / 2) * webSheetFormulaBesselI1(value) +
      (1 / value) * (1 + y * (0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * (-0.01919402 +
      y * (-0.00110404 + y * (-0.00004686)))))));
  }
  const y = 2 / value;
  return (Math.exp(-value) / Math.sqrt(value)) *
    (1.25331414 + y * (0.23498619 + y * (-0.03655620 + y * (0.01504268 + y * (-0.00780353 +
    y * (0.00325614 - y * 0.00068245))))));
}



function webSheetFormulaBesselI(x, order) {
  const n = webSheetFormulaBesselOrder(order);
  if (n instanceof CellError) return n;
  const value = Number(x);
  if (!Number.isFinite(value)) return webSheetFormulaError(ErrorType.VALUE);
  if (n === 0) return webSheetFormulaBesselI0(value);
  if (n === 1) return webSheetFormulaBesselI1(value);
  if (value === 0) return 0;
  const tox = 2 / Math.abs(value);
  let bip = 0;
  let bi = 1;
  let result = 0;
  const big = 1e100;
  const small = 1e-100;
  for (let j = 2 * (n + Math.floor(Math.sqrt(40 * n))); j > 0; j -= 1) {
    const bim = bip + j * tox * bi;
    bip = bi;
    bi = bim;
    if (Math.abs(bi) > big) {
      bi *= small;
      bip *= small;
      result *= small;
    }
    if (j === n) result = bip;
  }
  result *= webSheetFormulaBesselI0(value) / bi;
  return value < 0 && n % 2 ? -result : result;
}



function webSheetFormulaBesselJ(x, order) {
  const n = webSheetFormulaBesselOrder(order);
  if (n instanceof CellError) return n;
  const value = Number(x);
  if (!Number.isFinite(value)) return webSheetFormulaError(ErrorType.VALUE);
  if (n === 0) return webSheetFormulaBesselJ0(value);
  if (n === 1) return webSheetFormulaBesselJ1(value);
  if (value === 0) return 0;
  const ax = Math.abs(value);
  let previous = webSheetFormulaBesselJ0(ax);
  let current = webSheetFormulaBesselJ1(ax);
  for (let index = 1; index < n; index += 1) {
    const next = (2 * index / ax) * current - previous;
    previous = current;
    current = next;
  }
  return value < 0 && n % 2 ? -current : current;
}



function webSheetFormulaBesselK(x, order) {
  const n = webSheetFormulaBesselOrder(order);
  if (n instanceof CellError) return n;
  const value = Number(x);
  if (!Number.isFinite(value) || value <= 0) return webSheetFormulaError(ErrorType.NUM);
  if (n === 0) return webSheetFormulaBesselK0(value);
  if (n === 1) return webSheetFormulaBesselK1(value);
  let previous = webSheetFormulaBesselK0(value);
  let current = webSheetFormulaBesselK1(value);
  for (let index = 1; index < n; index += 1) {
    const next = previous + (2 * index / value) * current;
    previous = current;
    current = next;
  }
  return current;
}



function webSheetFormulaBesselY(x, order) {
  const n = webSheetFormulaBesselOrder(order);
  if (n instanceof CellError) return n;
  const value = Number(x);
  if (!Number.isFinite(value) || value <= 0) return webSheetFormulaError(ErrorType.NUM);
  if (n === 0) return webSheetFormulaBesselY0(value);
  if (n === 1) return webSheetFormulaBesselY1(value);
  let previous = webSheetFormulaBesselY0(value);
  let current = webSheetFormulaBesselY1(value);
  for (let index = 1; index < n; index += 1) {
    const next = (2 * index / value) * current - previous;
    previous = current;
    current = next;
  }
  return current;
}



function webSheetFormulaInfoN(value) {
  if (value instanceof CellError) return value;
  const raw = webSheetFormulaRawValue(webSheetFormulaScalar(value));
  if (raw instanceof CellError) return raw;
  if (typeof raw === "number") return raw;
  if (typeof raw === "boolean") return raw ? 1 : 0;
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) return dateToExcelSerial(raw);
  return 0;
}



function webSheetFormulaPercentile(sortedValues, k, exclusive = false) {
  const values = sortedValues.filter(Number.isFinite).sort((left, right) => left - right);
  const n = values.length;
  if (!n || k < 0 || k > 1 || (exclusive && (k <= 1 / (n + 1) || k >= n / (n + 1)))) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  if (n === 1) return values[0];
  const position = exclusive ? k * (n + 1) - 1 : k * (n - 1);
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) return values[lower];
  return values[lower] + (values[upper] - values[lower]) * (position - lower);
}



function webSheetFormulaAggregateValues(args, { ignoreErrors = false, numericOnly = true, countAll = false } = {}) {
  const values = [];
  const candidates = Array.isArray(args)
    ? args.flatMap((arg) => webSheetFormulaFlattenValues(arg))
    : webSheetFormulaFlattenValues(args);
  for (const value of candidates) {
    if (value instanceof CellError) {
      if (!ignoreErrors) return value;
      continue;
    }
    const raw = webSheetFormulaRawValue(value);
    if (raw instanceof CellError) {
      if (!ignoreErrors) return raw;
      continue;
    }
    if (countAll) {
      if (!webSheetFormulaIsEmptyValue(raw)) values.push(raw);
      continue;
    }
    if (!numericOnly) {
      values.push(raw);
      continue;
    }
    const number = webSheetFormulaNumber(value);
    if (number instanceof CellError) {
      if (!ignoreErrors) return number;
      continue;
    }
    if (Number.isFinite(number)) values.push(number);
  }
  return values;
}



function webSheetFormulaMedian(values) {
  const sorted = values.filter(Number.isFinite).sort((left, right) => left - right);
  if (!sorted.length) return webSheetFormulaError(ErrorType.NUM);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}



function webSheetFormulaModeSnglValue(values) {
  const counts = new Map();
  values.forEach((value) => counts.set(value, (counts.get(value) || 0) + 1));
  let best = null;
  let bestCount = 1;
  for (const value of values) {
    const count = counts.get(value) || 0;
    if (count > bestCount) {
      best = value;
      bestCount = count;
    }
  }
  return best == null ? webSheetFormulaError(ErrorType.NA) : best;
}



function webSheetFormulaModeMultValues(values) {
  const counts = new Map();
  values.forEach((value) => counts.set(value, (counts.get(value) || 0) + 1));
  const maxCount = Math.max(0, ...counts.values());
  if (maxCount <= 1) return webSheetFormulaError(ErrorType.NA);
  const seen = new Set();
  const modes = [];
  for (const value of values) {
    if ((counts.get(value) || 0) !== maxCount || seen.has(value)) continue;
    seen.add(value);
    modes.push(value);
  }
  return modes;
}



function webSheetFormulaRank(number, ref, order = 0, average = false) {
  const values = webSheetFormulaNumericValues([ref]);
  if (values instanceof CellError) return values;
  const target = webSheetFormulaNumber(number);
  if (!Number.isFinite(target)) return webSheetFormulaError(ErrorType.VALUE);
  const equals = values.filter((value) => value === target).length;
  if (!equals) return webSheetFormulaError(ErrorType.NA);
  const before = values.filter((value) => order ? value < target : value > target).length;
  return average ? before + (equals + 1) / 2 : before + 1;
}



function webSheetFormulaKurt(values) {
  const numbers = values.filter(Number.isFinite);
  const n = numbers.length;
  if (n < 4) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const variance = webSheetFormulaVariance(numbers, true);
  if (variance instanceof CellError) return variance;
  const sd = Math.sqrt(variance);
  if (!sd) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const mean = numbers.reduce((sum, value) => sum + value, 0) / n;
  const fourthMoment = numbers.reduce((sum, value) => sum + ((value - mean) / sd) ** 4, 0);
  return (n * (n + 1) * fourthMoment) / ((n - 1) * (n - 2) * (n - 3)) -
    (3 * (n - 1) * (n - 1)) / ((n - 2) * (n - 3));
}



function webSheetFormulaPercentRank(values, x, significance = 3, exclusive = false) {
  const sorted = values.filter(Number.isFinite).sort((left, right) => left - right);
  const n = sorted.length;
  const target = Number(x);
  const digits = Math.trunc(Number(significance ?? 3));
  if (!n || !Number.isFinite(target) || !Number.isFinite(digits) || digits < 1) return webSheetFormulaError(ErrorType.NUM);
  const min = sorted[0];
  const max = sorted[n - 1];
  if (exclusive ? target <= min || target >= max : target < min || target > max) return webSheetFormulaError(ErrorType.NA);
  let rank = null;
  for (let index = 0; index < n; index += 1) {
    if (sorted[index] === target) {
      rank = exclusive ? (index + 1) / (n + 1) : (n === 1 ? 0 : index / (n - 1));
      break;
    }
    if (sorted[index] > target) {
      const lowerIndex = index - 1;
      const upperIndex = index;
      const lowerRank = exclusive ? (lowerIndex + 1) / (n + 1) : lowerIndex / (n - 1);
      const upperRank = exclusive ? (upperIndex + 1) / (n + 1) : upperIndex / (n - 1);
      const fraction = (target - sorted[lowerIndex]) / (sorted[upperIndex] - sorted[lowerIndex]);
      rank = lowerRank + (upperRank - lowerRank) * fraction;
      break;
    }
  }
  if (rank == null) rank = exclusive ? n / (n + 1) : 1;
  const factor = 10 ** Math.min(15, digits);
  return Math.floor(rank * factor) / factor;
}



function webSheetFormulaVariance(values, sample = true) {
  const divisor = sample ? values.length - 1 : values.length;
  if (divisor <= 0) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  return values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / divisor;
}



function webSheetFormulaVarianceFromArgs(args, sample = true, includeLogicalText = false) {
  const values = includeLogicalText ? webSheetFormulaAValues(args) : webSheetFormulaNumericValues(args);
  return values instanceof CellError ? values : webSheetFormulaVariance(values, sample);
}



function webSheetFormulaStdevFromArgs(args, sample = true, includeLogicalText = false) {
  const variance = webSheetFormulaVarianceFromArgs(args, sample, includeLogicalText);
  return variance instanceof CellError ? variance : Math.sqrt(variance);
}



function webSheetFormulaAverageDeviation(values) {
  if (!values.length) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  return values.reduce((sum, value) => sum + Math.abs(value - mean), 0) / values.length;
}



function webSheetFormulaDevSq(values) {
  if (!values.length) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  return values.reduce((sum, value) => sum + (value - mean) ** 2, 0);
}



function webSheetFormulaGeoMean(values) {
  if (!values.length) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  if (values.some((value) => value <= 0)) return webSheetFormulaError(ErrorType.NUM);
  return Math.exp(values.reduce((sum, value) => sum + Math.log(value), 0) / values.length);
}



function webSheetFormulaHarMean(values) {
  if (!values.length) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  if (values.some((value) => value <= 0)) return webSheetFormulaError(ErrorType.NUM);
  return values.length / values.reduce((sum, value) => sum + 1 / value, 0);
}



function webSheetFormulaPairedNumericValues(left, right) {
  const leftValues = webSheetFormulaFlattenValues(left);
  const rightValues = webSheetFormulaFlattenValues(right);
  const pairs = [];
  const length = Math.min(leftValues.length, rightValues.length);
  for (let index = 0; index < length; index += 1) {
    const leftNumber = webSheetFormulaNumber(leftValues[index]);
    const rightNumber = webSheetFormulaNumber(rightValues[index]);
    if (leftNumber instanceof CellError) return leftNumber;
    if (rightNumber instanceof CellError) return rightNumber;
    if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) pairs.push([leftNumber, rightNumber]);
  }
  return pairs;
}



function webSheetFormulaCovariance(left, right, sample = true) {
  const pairs = webSheetFormulaPairedNumericValues(left, right);
  if (pairs instanceof CellError) return pairs;
  const divisor = sample ? pairs.length - 1 : pairs.length;
  if (divisor <= 0) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const meanLeft = pairs.reduce((sum, pair) => sum + pair[0], 0) / pairs.length;
  const meanRight = pairs.reduce((sum, pair) => sum + pair[1], 0) / pairs.length;
  return pairs.reduce((sum, pair) => sum + (pair[0] - meanLeft) * (pair[1] - meanRight), 0) / divisor;
}



function webSheetFormulaSumPairwiseSquares(left, right, mode) {
  const pairs = webSheetFormulaPairedNumericValues(left, right);
  if (pairs instanceof CellError) return pairs;
  return pairs.reduce((sum, [x, y]) => {
    if (mode === "x2my2") return sum + x * x - y * y;
    if (mode === "x2py2") return sum + x * x + y * y;
    return sum + (x - y) * (x - y);
  }, 0);
}



function webSheetFormulaCorrelation(left, right) {
  const pairs = webSheetFormulaPairedNumericValues(left, right);
  if (pairs instanceof CellError) return pairs;
  if (pairs.length < 2) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const meanLeft = pairs.reduce((sum, pair) => sum + pair[0], 0) / pairs.length;
  const meanRight = pairs.reduce((sum, pair) => sum + pair[1], 0) / pairs.length;
  const sumProducts = pairs.reduce((sum, pair) => sum + (pair[0] - meanLeft) * (pair[1] - meanRight), 0);
  const leftSquares = pairs.reduce((sum, pair) => sum + (pair[0] - meanLeft) ** 2, 0);
  const rightSquares = pairs.reduce((sum, pair) => sum + (pair[1] - meanRight) ** 2, 0);
  const denominator = Math.sqrt(leftSquares * rightSquares);
  return denominator ? sumProducts / denominator : webSheetFormulaError(ErrorType.DIV_BY_ZERO);
}



function webSheetFormulaSteyx(knownY, knownX) {
  const pairs = webSheetFormulaPairedNumericValues(knownY, knownX);
  if (pairs instanceof CellError) return pairs;
  if (pairs.length < 3) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const regression = webSheetFormulaRegression(knownY, knownX);
  if (regression instanceof CellError) return regression;
  const residualSquares = pairs.reduce((sum, pair) => {
    const y = pair[0];
    const x = pair[1];
    const predicted = regression.intercept + regression.slope * x;
    return sum + (y - predicted) ** 2;
  }, 0);
  return Math.sqrt(residualSquares / (pairs.length - 2));
}



function webSheetFormulaSkew(values, population = false) {
  if (population) {
    if (!values.length) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
    const variance = webSheetFormulaVariance(values, false);
    if (variance instanceof CellError) return variance;
    const sd = Math.sqrt(variance);
    if (!sd) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    return values.reduce((sum, value) => sum + ((value - mean) / sd) ** 3, 0) / values.length;
  }
  if (values.length < 3) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const variance = webSheetFormulaVariance(values, true);
  if (variance instanceof CellError) return variance;
  const sd = Math.sqrt(variance);
  if (!sd) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  return (values.length / ((values.length - 1) * (values.length - 2))) *
    values.reduce((sum, value) => sum + ((value - mean) / sd) ** 3, 0);
}



function webSheetFormulaErf(value) {
  const sign = value < 0 ? -1 : 1;
  const x = Math.abs(Number(value));
  const t = 1 / (1 + 0.3275911 * x);
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const y = 1 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
  return sign * y;
}



function webSheetFormulaNormalPdf(x, mean = 0, standardDeviation = 1) {
  if (standardDeviation <= 0) return webSheetFormulaError(ErrorType.NUM);
  const z = (x - mean) / standardDeviation;
  return Math.exp(-0.5 * z * z) / (standardDeviation * Math.sqrt(2 * Math.PI));
}



function webSheetFormulaNormalCdf(x, mean = 0, standardDeviation = 1) {
  if (standardDeviation <= 0) return webSheetFormulaError(ErrorType.NUM);
  return 0.5 * (1 + webSheetFormulaErf((x - mean) / (standardDeviation * Math.SQRT2)));
}



function webSheetFormulaInverseStandardNormal(probability) {
  const p = Number(probability);
  if (!(p > 0 && p < 1)) return webSheetFormulaError(ErrorType.NUM);
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const low = 0.02425;
  const high = 1 - low;
  let q;
  let result;
  if (p < low) {
    q = Math.sqrt(-2 * Math.log(p));
    result = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else if (p <= high) {
    q = p - 0.5;
    const r = q * q;
    result = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q /
      (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  } else {
    q = Math.sqrt(-2 * Math.log(1 - p));
    result = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  const error = webSheetFormulaNormalCdf(result, 0, 1) - p;
  const density = webSheetFormulaNormalPdf(result, 0, 1);
  return result - error / density;
}



function webSheetFormulaNormalInv(probability, mean = 0, standardDeviation = 1) {
  if (standardDeviation <= 0) return webSheetFormulaError(ErrorType.NUM);
  const z = webSheetFormulaInverseStandardNormal(probability);
  return z instanceof CellError ? z : mean + standardDeviation * z;
}



function webSheetFormulaFactorial(value) {
  const n = Math.trunc(Number(value));
  if (n < 0 || !Number.isFinite(n)) return NaN;
  let product = 1;
  for (let index = 2; index <= n; index += 1) product *= index;
  return product;
}



function webSheetFormulaCombination(n, k) {
  const total = Math.trunc(Number(n));
  const selected = Math.trunc(Number(k));
  if (selected < 0 || selected > total || total < 0) return NaN;
  const count = Math.min(selected, total - selected);
  let result = 1;
  for (let index = 1; index <= count; index += 1) {
    result *= (total - count + index) / index;
  }
  return result;
}



function webSheetFormulaFactorialResult(value) {
  const n = Math.trunc(Number(value));
  if (n < 0 || n > 170 || !Number.isFinite(n)) return webSheetFormulaError(ErrorType.NUM);
  return webSheetFormulaFactorial(n);
}



function webSheetFormulaDoubleFactorial(value) {
  const n = Math.trunc(Number(value));
  if (n < 0 || !Number.isFinite(n)) return webSheetFormulaError(ErrorType.NUM);
  let product = 1;
  for (let index = Math.max(1, n); index > 1; index -= 2) {
    product *= index;
    if (!Number.isFinite(product)) return webSheetFormulaError(ErrorType.NUM);
  }
  return product;
}



function webSheetFormulaCombinationResult(n, k) {
  const result = webSheetFormulaCombination(n, k);
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaCombinationA(n, k) {
  const total = Math.trunc(Number(n));
  const selected = Math.trunc(Number(k));
  if (total < 1 || selected < 0 || !Number.isFinite(total) || !Number.isFinite(selected)) return webSheetFormulaError(ErrorType.NUM);
  if (selected === 0) return 1;
  return webSheetFormulaCombinationResult(total + selected - 1, selected);
}



function webSheetFormulaPermutation(n, k) {
  const total = Math.trunc(Number(n));
  const selected = Math.trunc(Number(k));
  if (total < 0 || selected < 0 || selected > total || !Number.isFinite(total) || !Number.isFinite(selected)) return webSheetFormulaError(ErrorType.NUM);
  let result = 1;
  for (let index = total - selected + 1; index <= total; index += 1) {
    result *= index;
    if (!Number.isFinite(result)) return webSheetFormulaError(ErrorType.NUM);
  }
  return result;
}



function webSheetFormulaPermutationA(n, k) {
  const total = Math.trunc(Number(n));
  const selected = Math.trunc(Number(k));
  if (total < 0 || selected < 0 || !Number.isFinite(total) || !Number.isFinite(selected)) return webSheetFormulaError(ErrorType.NUM);
  const result = total ** selected;
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaGreatestCommonDivisor(values) {
  const gcd = (left, right) => {
    let a = Math.trunc(Math.abs(left));
    let b = Math.trunc(Math.abs(right));
    while (b) [a, b] = [b, a % b];
    return a;
  };
  if (values.some((value) => value < 0)) return webSheetFormulaError(ErrorType.NUM);
  return values.reduce((result, value) => gcd(result, value), 0);
}



function webSheetFormulaLeastCommonMultiple(values) {
  if (values.some((value) => value < 0)) return webSheetFormulaError(ErrorType.NUM);
  if (values.some((value) => value === 0)) return 0;
  const gcdValue = (left, right) => webSheetFormulaGreatestCommonDivisor([left, right]);
  return values.reduce((result, value) => {
    if (result instanceof CellError) return result;
    const divisor = gcdValue(result, value);
    if (divisor instanceof CellError) return divisor;
    const next = Math.abs(result * value) / divisor;
    return Number.isFinite(next) ? next : webSheetFormulaError(ErrorType.NUM);
  }, 1);
}



function webSheetFormulaMod(number, divisor) {
  if (divisor === 0) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  return number - divisor * Math.floor(number / divisor);
}



function webSheetFormulaTrigReciprocal(value, fn) {
  const denominator = fn(value);
  if (!Number.isFinite(denominator)) return webSheetFormulaError(ErrorType.NUM);
  if (Math.abs(denominator) < 1e-15) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  return 1 / denominator;
}



function webSheetFormulaAcot(value) {
  if (value === 0) return Math.PI / 2;
  const result = Math.atan(1 / value);
  return value < 0 ? result + Math.PI : result;
}



function webSheetFormulaPoissonDist(x, mean, cumulative = false) {
  const k = Math.trunc(Number(x));
  const lambda = Number(mean);
  if (k < 0 || lambda <= 0 || !Number.isFinite(k) || !Number.isFinite(lambda)) return webSheetFormulaError(ErrorType.NUM);
  const probability = (count) => Math.exp(-lambda) * (lambda ** count) / webSheetFormulaFactorial(count);
  if (!cumulative) return probability(k);
  let total = 0;
  for (let index = 0; index <= k; index += 1) total += probability(index);
  return total;
}



function webSheetFormulaExponDist(x, lambda, cumulative = false) {
  const value = Number(x);
  const rate = Number(lambda);
  if (value < 0 || rate <= 0 || !Number.isFinite(value) || !Number.isFinite(rate)) return webSheetFormulaError(ErrorType.NUM);
  return cumulative ? 1 - Math.exp(-rate * value) : rate * Math.exp(-rate * value);
}



function webSheetFormulaBinomDist(numberS, trials, probabilityS, cumulative = false) {
  const successes = Math.trunc(Number(numberS));
  const attempts = Math.trunc(Number(trials));
  const probability = Number(probabilityS);
  if (
    successes < 0 || attempts < 0 || successes > attempts ||
    probability < 0 || probability > 1 ||
    !Number.isFinite(successes) || !Number.isFinite(attempts) || !Number.isFinite(probability)
  ) return webSheetFormulaError(ErrorType.NUM);
  const mass = (count) => webSheetFormulaCombination(attempts, count) *
    (probability ** count) *
    ((1 - probability) ** (attempts - count));
  if (!cumulative) return mass(successes);
  let total = 0;
  for (let index = 0; index <= successes; index += 1) total += mass(index);
  return total;
}



function webSheetFormulaBinomDistRange(trials, probabilityS, numberS, numberS2 = WEBSHEET_FORMULA_NA_SENTINEL) {
  const attempts = Math.trunc(Number(trials));
  const probability = Number(probabilityS);
  const lower = Math.trunc(Number(numberS));
  const upper = numberS2 === WEBSHEET_FORMULA_NA_SENTINEL ? lower : Math.trunc(Number(numberS2));
  if (
    attempts < 0 || probability < 0 || probability > 1 || lower < 0 || upper < lower || upper > attempts ||
    !Number.isFinite(attempts) || !Number.isFinite(probability) || !Number.isFinite(lower) || !Number.isFinite(upper)
  ) return webSheetFormulaError(ErrorType.NUM);
  let total = 0;
  for (let count = lower; count <= upper; count += 1) {
    const value = webSheetFormulaBinomDist(count, attempts, probability, false);
    if (value instanceof CellError) return value;
    total += value;
  }
  return total;
}



function webSheetFormulaBinomInv(trials, probabilityS, alpha) {
  const attempts = Math.trunc(Number(trials));
  const probability = Number(probabilityS);
  const threshold = Number(alpha);
  if (
    attempts < 0 || probability < 0 || probability > 1 || threshold < 0 || threshold > 1 ||
    !Number.isFinite(attempts) || !Number.isFinite(probability) || !Number.isFinite(threshold)
  ) return webSheetFormulaError(ErrorType.NUM);
  let cumulative = 0;
  for (let successes = 0; successes <= attempts; successes += 1) {
    cumulative += webSheetFormulaBinomDist(successes, attempts, probability, false);
    if (cumulative >= threshold) return successes;
  }
  return attempts;
}



function webSheetFormulaLogGammaPositive(value) {
  const z = Number(value);
  if (!(z > 0) || !Number.isFinite(z)) return NaN;
  const coefficients = [
    676.5203681218851,
    -1259.1392167224028,
    771.3234287776531,
    -176.6150291621406,
    12.507343278686905,
    -0.13857109526572012,
    9.984369578019572e-6,
    1.5056327351493116e-7,
  ];
  if (z < 0.5) {
    return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * z)) - webSheetFormulaLogGammaPositive(1 - z);
  }
  let x = 0.9999999999998099;
  const adjusted = z - 1;
  coefficients.forEach((coefficient, index) => {
    x += coefficient / (adjusted + index + 1);
  });
  const t = adjusted + coefficients.length - 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (adjusted + 0.5) * Math.log(t) - t + Math.log(x);
}



function webSheetFormulaGamma(value) {
  const z = Number(value);
  if (!Number.isFinite(z) || z === 0 || (z < 0 && Math.trunc(z) === z)) return webSheetFormulaError(ErrorType.NUM);
  const result = z < 0.5
    ? Math.PI / (Math.sin(Math.PI * z) * Math.exp(webSheetFormulaLogGammaPositive(1 - z)))
    : Math.exp(webSheetFormulaLogGammaPositive(z));
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaGammaSeriesP(a, x) {
  const maxIterations = 200;
  const epsilon = 1e-12;
  let ap = a;
  let sum = 1 / a;
  let del = sum;
  for (let index = 1; index <= maxIterations; index += 1) {
    ap += 1;
    del *= x / ap;
    sum += del;
    if (Math.abs(del) < Math.abs(sum) * epsilon) break;
  }
  return sum * Math.exp(-x + a * Math.log(x) - webSheetFormulaLogGammaPositive(a));
}



function webSheetFormulaGammaContinuedFractionQ(a, x) {
  const maxIterations = 200;
  const epsilon = 1e-12;
  const fpmin = 1e-300;
  let b = x + 1 - a;
  let c = 1 / fpmin;
  let d = Math.abs(b) < fpmin ? fpmin : b;
  d = 1 / d;
  let h = d;
  for (let index = 1; index <= maxIterations; index += 1) {
    const an = -index * (index - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < fpmin) d = fpmin;
    c = b + an / c;
    if (Math.abs(c) < fpmin) c = fpmin;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < epsilon) break;
  }
  return Math.exp(-x + a * Math.log(x) - webSheetFormulaLogGammaPositive(a)) * h;
}



function webSheetFormulaRegularizedGammaP(a, x) {
  const shape = Number(a);
  const value = Number(x);
  if (!(shape > 0) || value < 0 || !Number.isFinite(shape) || !Number.isFinite(value)) return NaN;
  if (value === 0) return 0;
  const result = value < shape + 1
    ? webSheetFormulaGammaSeriesP(shape, value)
    : 1 - webSheetFormulaGammaContinuedFractionQ(shape, value);
  return Math.min(1, Math.max(0, result));
}



function webSheetFormulaRegularizedBeta(x, alpha, beta) {
  const value = Number(x);
  const a = Number(alpha);
  const b = Number(beta);
  if (!(a > 0) || !(b > 0) || value < 0 || value > 1 || !Number.isFinite(a) || !Number.isFinite(b) || !Number.isFinite(value)) return NaN;
  if (value <= 0) return 0;
  if (value >= 1) return 1;
  return dataAnalysisRegularizedBeta(value, a, b);
}



function webSheetFormulaInverseByBisection(probability, cumulativeFn, low = 0, high = 1) {
  const target = Number(probability);
  if (!(target >= 0 && target <= 1) || !Number.isFinite(target)) return NaN;
  if (target === 0) return low;
  let upper = high;
  let guard = 0;
  while (cumulativeFn(upper) < target && upper < 1e12 && guard < 120) {
    upper *= 2;
    guard += 1;
  }
  if (cumulativeFn(upper) < target) return NaN;
  let lower = low;
  for (let index = 0; index < 90; index += 1) {
    const mid = (lower + upper) / 2;
    if (cumulativeFn(mid) < target) lower = mid;
    else upper = mid;
  }
  return upper;
}



function webSheetFormulaGammaDist(x, alpha, beta, cumulative = false) {
  const value = Number(x);
  const shape = Number(alpha);
  const scale = Number(beta);
  if (value < 0 || !(shape > 0) || !(scale > 0) || !Number.isFinite(value) || !Number.isFinite(shape) || !Number.isFinite(scale)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  if (cumulative) return webSheetFormulaRegularizedGammaP(shape, value / scale);
  if (value === 0) {
    if (shape === 1) return 1 / scale;
    return shape > 1 ? 0 : webSheetFormulaError(ErrorType.NUM);
  }
  const logPdf = (shape - 1) * Math.log(value / scale) - (value / scale) - webSheetFormulaLogGammaPositive(shape) - Math.log(scale);
  const result = Math.exp(logPdf);
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaGammaInv(probability, alpha, beta) {
  const p = Number(probability);
  const shape = Number(alpha);
  const scale = Number(beta);
  if (!(p >= 0 && p < 1) || !(shape > 0) || !(scale > 0) || !Number.isFinite(p) || !Number.isFinite(shape) || !Number.isFinite(scale)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const result = webSheetFormulaInverseByBisection(p, (value) => webSheetFormulaRegularizedGammaP(shape, value / scale), 0, Math.max(scale, shape * scale));
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaBetaDist(x, alpha, beta, cumulative = true, lower = 0, upper = 1) {
  const value = Number(x);
  const a = Number(alpha);
  const b = Number(beta);
  const min = Number(lower);
  const max = Number(upper);
  if (!(a > 0) || !(b > 0) || !(max > min) || value < min || value > max ||
    !Number.isFinite(value) || !Number.isFinite(a) || !Number.isFinite(b) || !Number.isFinite(min) || !Number.isFinite(max)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const scaled = (value - min) / (max - min);
  if (cumulative) return webSheetFormulaRegularizedBeta(scaled, a, b);
  if (scaled === 0 || scaled === 1) {
    if ((scaled === 0 && a < 1) || (scaled === 1 && b < 1)) return webSheetFormulaError(ErrorType.NUM);
    if (scaled === 0 && a === 1) return b / (max - min);
    if (scaled === 1 && b === 1) return a / (max - min);
    return 0;
  }
  const logBeta = webSheetFormulaLogGammaPositive(a) + webSheetFormulaLogGammaPositive(b) - webSheetFormulaLogGammaPositive(a + b);
  const logPdf = (a - 1) * Math.log(scaled) + (b - 1) * Math.log(1 - scaled) - logBeta - Math.log(max - min);
  const result = Math.exp(logPdf);
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaBetaInv(probability, alpha, beta, lower = 0, upper = 1) {
  const p = Number(probability);
  const a = Number(alpha);
  const b = Number(beta);
  const min = Number(lower);
  const max = Number(upper);
  if (!(p >= 0 && p <= 1) || !(a > 0) || !(b > 0) || !(max > min) ||
    !Number.isFinite(p) || !Number.isFinite(a) || !Number.isFinite(b) || !Number.isFinite(min) || !Number.isFinite(max)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  if (p === 0) return min;
  if (p === 1) return max;
  const scaled = webSheetFormulaInverseByBisection(p, (value) => webSheetFormulaRegularizedBeta(value, a, b), 0, 1);
  return Number.isFinite(scaled) ? min + scaled * (max - min) : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaChiSqDist(x, degFreedom, cumulative = false) {
  const value = Number(x);
  const df = Number(degFreedom);
  if (value < 0 || !(df > 0) || !Number.isFinite(value) || !Number.isFinite(df)) return webSheetFormulaError(ErrorType.NUM);
  return cumulative ? webSheetFormulaRegularizedGammaP(df / 2, value / 2) : webSheetFormulaGammaDist(value, df / 2, 2, false);
}



function webSheetFormulaChiSqInv(probability, degFreedom, rightTail = false) {
  const p = Number(probability);
  const df = Number(degFreedom);
  if (!(p >= 0 && p < 1) || !(df > 0) || !Number.isFinite(p) || !Number.isFinite(df)) return webSheetFormulaError(ErrorType.NUM);
  const target = rightTail ? 1 - p : p;
  const result = webSheetFormulaInverseByBisection(target, (value) => webSheetFormulaRegularizedGammaP(df / 2, value / 2), 0, Math.max(1, df));
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaWeibullDist(x, alpha, beta, cumulative = false) {
  const value = Number(x);
  const shape = Number(alpha);
  const scale = Number(beta);
  if (value < 0 || !(shape > 0) || !(scale > 0) || !Number.isFinite(value) || !Number.isFinite(shape) || !Number.isFinite(scale)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const powered = (value / scale) ** shape;
  return cumulative ? 1 - Math.exp(-powered) : (shape / scale) * ((value / scale) ** (shape - 1)) * Math.exp(-powered);
}



function webSheetFormulaLogCombination(n, k) {
  const total = Math.trunc(Number(n));
  const selected = Math.trunc(Number(k));
  if (total < 0 || selected < 0 || selected > total || !Number.isFinite(total) || !Number.isFinite(selected)) return -Infinity;
  return webSheetFormulaLogGammaPositive(total + 1) - webSheetFormulaLogGammaPositive(selected + 1) - webSheetFormulaLogGammaPositive(total - selected + 1);
}



function webSheetFormulaHypergeomDist(sampleS, numberSample, populationS, numberPop, cumulative = false) {
  const k = Math.trunc(Number(sampleS));
  const n = Math.trunc(Number(numberSample));
  const good = Math.trunc(Number(populationS));
  const total = Math.trunc(Number(numberPop));
  if (
    k < 0 || n < 0 || good < 0 || total < 0 || k > n || k > good || n > total || good > total ||
    !Number.isFinite(k) || !Number.isFinite(n) || !Number.isFinite(good) || !Number.isFinite(total)
  ) return webSheetFormulaError(ErrorType.NUM);
  const mass = (successes) => {
    const logProbability = webSheetFormulaLogCombination(good, successes) +
      webSheetFormulaLogCombination(total - good, n - successes) -
      webSheetFormulaLogCombination(total, n);
    return Number.isFinite(logProbability) ? Math.exp(logProbability) : 0;
  };
  if (!cumulative) return mass(k);
  const lower = Math.max(0, n - (total - good));
  const upper = Math.min(k, n, good);
  let sum = 0;
  for (let index = lower; index <= upper; index += 1) sum += mass(index);
  return sum;
}



function webSheetFormulaConfidenceNorm(alpha, standardDev, size) {
  const significance = Number(alpha);
  const sd = Number(standardDev);
  const count = Math.trunc(Number(size));
  if (!(significance > 0 && significance < 1) || !(sd > 0) || count < 1 ||
    !Number.isFinite(significance) || !Number.isFinite(sd) || !Number.isFinite(count)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const z = webSheetFormulaInverseStandardNormal(1 - significance / 2);
  return z instanceof CellError ? z : z * sd / Math.sqrt(count);
}



function webSheetFormulaConfidenceT(alpha, standardDev, size) {
  const significance = Number(alpha);
  const sd = Number(standardDev);
  const count = Math.trunc(Number(size));
  if (!(significance > 0 && significance < 1) || !(sd > 0) || count < 2 ||
    !Number.isFinite(significance) || !Number.isFinite(sd) || !Number.isFinite(count)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const t = dataAnalysisStudentTInverseTwoTail(significance, count - 1);
  return Number.isFinite(t) ? t * sd / Math.sqrt(count) : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaFDist(x, degFreedom1, degFreedom2, cumulative = false) {
  const value = Number(x);
  const df1 = Number(degFreedom1);
  const df2 = Number(degFreedom2);
  if (value < 0 || !(df1 > 0) || !(df2 > 0) || !Number.isFinite(value) || !Number.isFinite(df1) || !Number.isFinite(df2)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  if (value === 0) return cumulative ? 0 : df1 === 2 ? 1 : 0;
  const z = (df1 * value) / (df1 * value + df2);
  if (cumulative) return webSheetFormulaRegularizedBeta(z, df1 / 2, df2 / 2);
  const logPdf = (df1 / 2) * Math.log(df1 / df2) +
    (df1 / 2 - 1) * Math.log(value) -
    webSheetFormulaLogGammaPositive(df1 / 2) -
    webSheetFormulaLogGammaPositive(df2 / 2) +
    webSheetFormulaLogGammaPositive((df1 + df2) / 2) -
    ((df1 + df2) / 2) * Math.log(1 + (df1 * value) / df2);
  const result = Math.exp(logPdf);
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaFInv(probability, degFreedom1, degFreedom2, rightTail = false) {
  const p = Number(probability);
  const df1 = Number(degFreedom1);
  const df2 = Number(degFreedom2);
  if (!(p >= 0 && p < 1) || !(df1 > 0) || !(df2 > 0) || !Number.isFinite(p) || !Number.isFinite(df1) || !Number.isFinite(df2)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const target = rightTail ? 1 - p : p;
  const result = webSheetFormulaInverseByBisection(target, (value) => {
    const cdf = webSheetFormulaFDist(value, df1, df2, true);
    return cdf instanceof CellError ? NaN : cdf;
  }, 0, 1);
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaStudentTPdf(x, degFreedom) {
  const value = Number(x);
  const df = Number(degFreedom);
  if (!(df > 0) || !Number.isFinite(value) || !Number.isFinite(df)) return webSheetFormulaError(ErrorType.NUM);
  const logPdf = webSheetFormulaLogGammaPositive((df + 1) / 2) -
    webSheetFormulaLogGammaPositive(df / 2) -
    0.5 * Math.log(df * Math.PI) -
    ((df + 1) / 2) * Math.log(1 + value * value / df);
  const result = Math.exp(logPdf);
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaStudentTCdf(x, degFreedom) {
  const value = Number(x);
  const df = Number(degFreedom);
  if (!(df > 0) || !Number.isFinite(value) || !Number.isFinite(df)) return webSheetFormulaError(ErrorType.NUM);
  if (value === 0) return 0.5;
  const beta = webSheetFormulaRegularizedBeta(df / (df + value * value), df / 2, 0.5);
  if (!Number.isFinite(beta)) return webSheetFormulaError(ErrorType.NUM);
  return value > 0 ? 1 - beta / 2 : beta / 2;
}



function webSheetFormulaStudentTInv(probability, degFreedom) {
  const p = Number(probability);
  const df = Number(degFreedom);
  if (!(p > 0 && p < 1) || !(df > 0) || !Number.isFinite(p) || !Number.isFinite(df)) return webSheetFormulaError(ErrorType.NUM);
  if (p === 0.5) return 0;
  const twoTail = 2 * Math.min(p, 1 - p);
  const magnitude = dataAnalysisStudentTInverseTwoTail(twoTail, df);
  return Number.isFinite(magnitude) ? (p < 0.5 ? -magnitude : magnitude) : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaStudentTDistTwoTail(x, degFreedom) {
  const value = Number(x);
  const df = Number(degFreedom);
  if (value < 0 || !(df > 0) || !Number.isFinite(value) || !Number.isFinite(df)) return webSheetFormulaError(ErrorType.NUM);
  const result = dataAnalysisStudentTTwoTail(value, df);
  return Number.isFinite(result) ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaMatrixValues(value) {
  if (value instanceof SimpleRangeValue) return value.rawData();
  if (Array.isArray(value)) return value.map((row) => (Array.isArray(row) ? row : [row]));
  return [[value]];
}



function webSheetFormulaZTest(array, x, sigma = WEBSHEET_FORMULA_NA_SENTINEL) {
  const values = webSheetFormulaNumericValues([array]);
  if (values instanceof CellError) return values;
  if (!values.length) return webSheetFormulaError(ErrorType.NA);
  const threshold = Number(webSheetFormulaRawValue(x));
  if (!Number.isFinite(threshold)) return webSheetFormulaError(ErrorType.VALUE);
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  const standardDeviation = sigma === WEBSHEET_FORMULA_NA_SENTINEL ? Math.sqrt(webSheetFormulaVariance(values, true)) : Number(webSheetFormulaRawValue(sigma));
  if (!Number.isFinite(standardDeviation) || standardDeviation <= 0) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const z = (mean - threshold) / (standardDeviation / Math.sqrt(values.length));
  return 1 - webSheetFormulaNormalCdf(z, 0, 1);
}



function webSheetFormulaFTest(array1, array2) {
  const left = webSheetFormulaNumericValues([array1]);
  const right = webSheetFormulaNumericValues([array2]);
  if (left instanceof CellError) return left;
  if (right instanceof CellError) return right;
  if (left.length < 2 || right.length < 2) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const leftVariance = webSheetFormulaVariance(left, true);
  const rightVariance = webSheetFormulaVariance(right, true);
  if (leftVariance instanceof CellError) return leftVariance;
  if (rightVariance instanceof CellError) return rightVariance;
  if (leftVariance <= 0 || rightVariance <= 0) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const useLeftAsNumerator = leftVariance >= rightVariance;
  const f = useLeftAsNumerator ? leftVariance / rightVariance : rightVariance / leftVariance;
  const df1 = useLeftAsNumerator ? left.length - 1 : right.length - 1;
  const df2 = useLeftAsNumerator ? right.length - 1 : left.length - 1;
  const cumulative = webSheetFormulaFDist(f, df1, df2, true);
  if (cumulative instanceof CellError) return cumulative;
  return Math.min(1, Math.max(0, 2 * (1 - cumulative)));
}



function webSheetFormulaChiSqTest(actualRange, expectedRange) {
  const actual = webSheetFormulaMatrixValues(actualRange);
  const expected = webSheetFormulaMatrixValues(expectedRange);
  if (!actual.length || !expected.length || actual.length !== expected.length) return webSheetFormulaError(ErrorType.NA);
  const rowCount = actual.length;
  const colCount = actual[0]?.length || 0;
  if (!colCount || rowCount < 2 || colCount < 2) return webSheetFormulaError(ErrorType.NA);
  let statistic = 0;
  for (let row = 0; row < rowCount; row += 1) {
    if (!Array.isArray(actual[row]) || !Array.isArray(expected[row]) || actual[row].length !== colCount || expected[row].length !== colCount) {
      return webSheetFormulaError(ErrorType.NA);
    }
    for (let col = 0; col < colCount; col += 1) {
      const observed = webSheetFormulaNumber(actual[row][col]);
      const estimated = webSheetFormulaNumber(expected[row][col]);
      if (observed instanceof CellError) return observed;
      if (estimated instanceof CellError) return estimated;
      if (!Number.isFinite(observed) || !Number.isFinite(estimated) || estimated <= 0) return webSheetFormulaError(ErrorType.NUM);
      statistic += ((observed - estimated) ** 2) / estimated;
    }
  }
  const df = (rowCount - 1) * (colCount - 1);
  const cumulative = webSheetFormulaChiSqDist(statistic, df, true);
  return cumulative instanceof CellError ? cumulative : 1 - cumulative;
}



function webSheetFormulaTTest(array1, array2, tails, type) {
  const tailCount = Math.trunc(Number(webSheetFormulaRawValue(tails)));
  const testType = Math.trunc(Number(webSheetFormulaRawValue(type)));
  if (![1, 2].includes(tailCount) || ![1, 2, 3].includes(testType)) return webSheetFormulaError(ErrorType.NUM);
  let tStatistic;
  let degreesFreedom;
  if (testType === 1) {
    const pairs = webSheetFormulaPairedNumericValues(array1, array2);
    if (pairs instanceof CellError) return pairs;
    if (pairs.length < 2) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
    const differences = pairs.map(([left, right]) => left - right);
    const variance = webSheetFormulaVariance(differences, true);
    if (variance instanceof CellError) return variance;
    const sd = Math.sqrt(variance);
    if (!sd) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
    const mean = differences.reduce((sum, value) => sum + value, 0) / differences.length;
    tStatistic = mean / (sd / Math.sqrt(differences.length));
    degreesFreedom = differences.length - 1;
  } else {
    const left = webSheetFormulaNumericValues([array1]);
    const right = webSheetFormulaNumericValues([array2]);
    if (left instanceof CellError) return left;
    if (right instanceof CellError) return right;
    if (left.length < 2 || right.length < 2) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
    const meanLeft = left.reduce((sum, value) => sum + value, 0) / left.length;
    const meanRight = right.reduce((sum, value) => sum + value, 0) / right.length;
    const varianceLeft = webSheetFormulaVariance(left, true);
    const varianceRight = webSheetFormulaVariance(right, true);
    if (varianceLeft instanceof CellError) return varianceLeft;
    if (varianceRight instanceof CellError) return varianceRight;
    if (testType === 2) {
      degreesFreedom = left.length + right.length - 2;
      const pooled = (((left.length - 1) * varianceLeft) + ((right.length - 1) * varianceRight)) / degreesFreedom;
      const standardError = Math.sqrt(pooled * (1 / left.length + 1 / right.length));
      if (!standardError) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
      tStatistic = (meanLeft - meanRight) / standardError;
    } else {
      const leftTerm = varianceLeft / left.length;
      const rightTerm = varianceRight / right.length;
      const standardError = Math.sqrt(leftTerm + rightTerm);
      if (!standardError) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
      tStatistic = (meanLeft - meanRight) / standardError;
      degreesFreedom = ((leftTerm + rightTerm) ** 2) /
        ((leftTerm ** 2) / (left.length - 1) + (rightTerm ** 2) / (right.length - 1));
    }
  }
  const cumulative = webSheetFormulaStudentTCdf(Math.abs(tStatistic), degreesFreedom);
  if (cumulative instanceof CellError) return cumulative;
  const oneTail = 1 - cumulative;
  return Math.min(1, Math.max(0, tailCount === 1 ? oneTail : oneTail * 2));
}



function webSheetFormulaLogNormDist(x, mean, standardDev, cumulative = true) {
  const value = Number(x);
  const mu = Number(mean);
  const sigma = Number(standardDev);
  if (!(value > 0) || !(sigma > 0) || !Number.isFinite(value) || !Number.isFinite(mu) || !Number.isFinite(sigma)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const z = (Math.log(value) - mu) / sigma;
  return cumulative ? webSheetFormulaNormalCdf(Math.log(value), mu, sigma) : Math.exp(-0.5 * z * z) / (value * sigma * Math.sqrt(2 * Math.PI));
}



function webSheetFormulaLogNormInv(probability, mean, standardDev) {
  const p = Number(probability);
  const mu = Number(mean);
  const sigma = Number(standardDev);
  if (!(p > 0 && p < 1) || !(sigma > 0) || !Number.isFinite(p) || !Number.isFinite(mu) || !Number.isFinite(sigma)) {
    return webSheetFormulaError(ErrorType.NUM);
  }
  const normal = webSheetFormulaNormalInv(p, mu, sigma);
  return normal instanceof CellError ? normal : Math.exp(normal);
}



function webSheetFormulaNegbinomDist(numberF, numberS, probabilityS, cumulative = false) {
  const failures = Math.trunc(Number(numberF));
  const successes = Math.trunc(Number(numberS));
  const probability = Number(probabilityS);
  if (
    failures < 0 || successes < 1 || probability < 0 || probability > 1 ||
    !Number.isFinite(failures) || !Number.isFinite(successes) || !Number.isFinite(probability)
  ) return webSheetFormulaError(ErrorType.NUM);
  const mass = (count) => {
    const logProbability = webSheetFormulaLogCombination(count + successes - 1, count) +
      successes * Math.log(probability) +
      count * Math.log(1 - probability);
    return Number.isFinite(logProbability) ? Math.exp(logProbability) : probability === 1 && count === 0 ? 1 : 0;
  };
  if (!cumulative) return mass(failures);
  let sum = 0;
  for (let index = 0; index <= failures; index += 1) sum += mass(index);
  return sum;
}



function webSheetFormulaQuartile(values, quart, exclusive = false) {
  const q = Math.trunc(Number(quart));
  if (exclusive && (q < 1 || q > 3)) return webSheetFormulaError(ErrorType.NUM);
  if (!exclusive && (q < 0 || q > 4)) return webSheetFormulaError(ErrorType.NUM);
  return webSheetFormulaPercentile(values, q / 4, exclusive);
}



function webSheetFormulaKthValue(values, k, largest = true) {
  const index = Math.trunc(Number(k));
  const sorted = values.filter(Number.isFinite).sort((left, right) => largest ? right - left : left - right);
  if (index < 1 || index > sorted.length) return webSheetFormulaError(ErrorType.NUM);
  return sorted[index - 1];
}



function webSheetFormulaAggregate(functionNum, options, args) {
  const code = Math.trunc(Number(functionNum));
  const option = Math.trunc(Number(options));
  if (!Number.isFinite(code) || !Number.isFinite(option) || code < 1 || code > 19 || option < 0 || option > 7) {
    return webSheetFormulaError(ErrorType.VALUE);
  }
  const ignoreErrors = [2, 3, 6, 7].includes(option);
  const valuesArg = code >= 14 ? args[0] : args;
  const values = webSheetFormulaAggregateValues(valuesArg, { ignoreErrors, numericOnly: code !== 3, countAll: code === 3 });
  if (values instanceof CellError) return values;
  if (code === 2) return values.length;
  if (code === 3) return values.length;
  const numbers = values.filter(Number.isFinite);
  if (!numbers.length) {
    if ([4, 5, 6, 9].includes(code)) return code === 6 ? 0 : 0;
    return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  }
  switch (code) {
    case 1:
      return numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
    case 4:
      return Math.max(...numbers);
    case 5:
      return Math.min(...numbers);
    case 6:
      return numbers.reduce((product, value) => product * value, 1);
    case 7:
      return Math.sqrt(webSheetFormulaVariance(numbers, true));
    case 8:
      return Math.sqrt(webSheetFormulaVariance(numbers, false));
    case 9:
      return numbers.reduce((sum, value) => sum + value, 0);
    case 10:
      return webSheetFormulaVariance(numbers, true);
    case 11:
      return webSheetFormulaVariance(numbers, false);
    case 12:
      return webSheetFormulaMedian(numbers);
    case 13:
      return webSheetFormulaModeSnglValue(numbers);
    case 14:
      return webSheetFormulaKthValue(numbers, args[1], true);
    case 15:
      return webSheetFormulaKthValue(numbers, args[1], false);
    case 16:
      return webSheetFormulaPercentile(numbers, Number(args[1]), false);
    case 17:
      return webSheetFormulaQuartile(numbers, args[1], false);
    case 18:
      return webSheetFormulaPercentile(numbers, Number(args[1]), true);
    case 19:
      return webSheetFormulaQuartile(numbers, args[1], true);
    default:
      return webSheetFormulaError(ErrorType.VALUE);
  }
}



function webSheetFormulaRegression(valuesY, valuesX) {
  const fit = webSheetFormulaMultipleRegression(valuesY, valuesX, { constant: true });
  if (fit instanceof CellError) return fit;
  return { slope: fit.coefficients[0], intercept: fit.intercept };
}



function webSheetFormulaTrimMean(array, percent) {
  const values = webSheetFormulaNumericValues([array]);
  if (values instanceof CellError) return values;
  values.sort((left, right) => left - right);
  const trimPercent = Number(percent);
  if (!values.length) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  if (!(trimPercent >= 0 && trimPercent < 1) || !Number.isFinite(trimPercent)) return webSheetFormulaError(ErrorType.NUM);
  const trimEachSide = Math.floor((values.length * trimPercent) / 2);
  const trimmed = values.slice(trimEachSide, values.length - trimEachSide);
  if (!trimmed.length) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  return trimmed.reduce((sum, value) => sum + value, 0) / trimmed.length;
}



function webSheetFormulaProb(xRange, probRange, lowerLimit, upperLimit = WEBSHEET_FORMULA_NA_SENTINEL) {
  const xValues = webSheetFormulaFlattenValues(xRange);
  const probabilities = webSheetFormulaFlattenValues(probRange);
  if (xValues.length !== probabilities.length) return webSheetFormulaError(ErrorType.NA);
  const lower = webSheetFormulaNumber(lowerLimit);
  const upper = upperLimit === WEBSHEET_FORMULA_NA_SENTINEL ? lower : webSheetFormulaNumber(upperLimit);
  if (lower instanceof CellError) return lower;
  if (upper instanceof CellError) return upper;
  if (!Number.isFinite(lower) || !Number.isFinite(upper)) return webSheetFormulaError(ErrorType.VALUE);
  let probabilitySum = 0;
  let result = 0;
  for (let index = 0; index < xValues.length; index += 1) {
    const x = webSheetFormulaNumber(xValues[index]);
    const probability = webSheetFormulaNumber(probabilities[index]);
    if (x instanceof CellError) return x;
    if (probability instanceof CellError) return probability;
    if (!Number.isFinite(x) || !Number.isFinite(probability) || probability < 0 || probability > 1) return webSheetFormulaError(ErrorType.NUM);
    probabilitySum += probability;
    if (x >= lower && x <= upper) result += probability;
  }
  return Math.abs(probabilitySum - 1) <= 1e-7 ? result : webSheetFormulaError(ErrorType.NUM);
}



function webSheetFormulaRegressionNumericVector(value) {
  const matrix = webSheetFormulaRegressionRawMatrix(value);
  const { height, width } = webSheetFormulaMatrixDimensions(matrix);
  const orientation = height === 1 && width > 1 ? "row" : "column";
  const values = [];
  for (const item of matrix.flatMap((row) => row)) {
    const number = webSheetFormulaNumber(item);
    if (number instanceof CellError) return number;
    values.push(Number.isFinite(number) ? number : null);
  }
  return { values, height, width, orientation };
}



function webSheetFormulaRegressionRawMatrix(value) {
  if (value instanceof SimpleRangeValue) return value.rawData().map((row) => row.map((item) => webSheetFormulaRawValue(item)));
  if (Array.isArray(value)) {
    if (value.some(Array.isArray)) {
      return value.map((row) => (Array.isArray(row) ? row : [row]).map((item) => webSheetFormulaRawValue(item)));
    }
    return [value.map((item) => webSheetFormulaRawValue(item))];
  }
  return [[webSheetFormulaRawValue(value)]];
}



function webSheetFormulaRegressionNumericMatrix(value) {
  const matrix = webSheetFormulaRegressionRawMatrix(value);
  const { height, width } = webSheetFormulaMatrixDimensions(matrix);
  if (!height || !width) return webSheetFormulaError(ErrorType.VALUE);
  const rows = [];
  for (let rowIndex = 0; rowIndex < height; rowIndex += 1) {
    const row = matrix[rowIndex] || [];
    if (row.length !== width) return webSheetFormulaError(ErrorType.VALUE);
    const numericRow = [];
    for (let colIndex = 0; colIndex < width; colIndex += 1) {
      const number = webSheetFormulaNumber(row[colIndex]);
      if (number instanceof CellError) return number;
      numericRow.push(Number.isFinite(number) ? number : null);
    }
    rows.push(numericRow);
  }
  return { rows, height, width };
}



function webSheetFormulaSequenceXRows(length) {
  return Array.from({ length }, (_unused, index) => [index + 1]);
}



function webSheetFormulaKnownXRows(knownX, yLength, yOrientation = "column") {
  if (knownX === WEBSHEET_FORMULA_NA_SENTINEL || knownX == null) {
    return {
      rows: webSheetFormulaSequenceXRows(yLength),
      featureCount: 1,
      outputShape: yOrientation,
      sourceHeight: yOrientation === "row" ? 1 : yLength,
      sourceWidth: yOrientation === "row" ? yLength : 1,
    };
  }
  const matrix = webSheetFormulaRegressionNumericMatrix(knownX);
  if (matrix instanceof CellError) return matrix;
  const { rows, height, width } = matrix;
  if (height === yLength) {
    return { rows: rows.map((row) => row.slice()), featureCount: width, outputShape: "column", sourceHeight: height, sourceWidth: width };
  }
  if (width === yLength) {
    return {
      rows: Array.from({ length: width }, (_unused, colIndex) => rows.map((row) => row[colIndex])),
      featureCount: height,
      outputShape: "row",
      sourceHeight: height,
      sourceWidth: width,
    };
  }
  if (height * width === yLength) {
    return {
      rows: rows.flatMap((row) => row).map((item) => [item]),
      featureCount: 1,
      outputShape: yOrientation,
      sourceHeight: yOrientation === "row" ? 1 : yLength,
      sourceWidth: yOrientation === "row" ? yLength : 1,
    };
  }
  return webSheetFormulaError(ErrorType.VALUE);
}



function webSheetFormulaTransposeMatrix(matrix) {
  const height = matrix.length;
  const width = Math.max(0, ...matrix.map((row) => row.length));
  return Array.from({ length: width }, (_unused, colIndex) =>
    Array.from({ length: height }, (_unusedRow, rowIndex) => matrix[rowIndex][colIndex]),
  );
}



function webSheetFormulaMatrixVectorMultiply(matrix, vector) {
  return matrix.map((row) => row.reduce((sum, value, index) => sum + value * vector[index], 0));
}



function webSheetFormulaMultipleRegression(valuesY, valuesX = WEBSHEET_FORMULA_NA_SENTINEL, options = {}) {
  const constant = options.constant !== false;
  const logY = Boolean(options.logY);
  const yVector = webSheetFormulaRegressionNumericVector(valuesY);
  if (yVector instanceof CellError) return yVector;
  const xSpec = webSheetFormulaKnownXRows(valuesX, yVector.values.length, yVector.orientation);
  if (xSpec instanceof CellError) return xSpec;
  const observations = [];
  for (let index = 0; index < yVector.values.length; index += 1) {
    const y = yVector.values[index];
    const xRow = xSpec.rows[index] || [];
    const hasCompleteX = xRow.length === xSpec.featureCount && xRow.every(Number.isFinite);
    if (!Number.isFinite(y) || !hasCompleteX) continue;
    if (logY && y <= 0) return webSheetFormulaError(ErrorType.NUM);
    observations.push({ y: logY ? Math.log(y) : y, x: xRow.slice() });
  }
  const parameterCount = xSpec.featureCount + (constant ? 1 : 0);
  if (!xSpec.featureCount || observations.length < parameterCount) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  const design = observations.map((observation) => constant ? [...observation.x, 1] : observation.x.slice());
  const y = observations.map((observation) => observation.y);
  const transposed = webSheetFormulaTransposeMatrix(design);
  const xtx = transposed.map((row) => transposed.map((otherRow) => row.reduce((sum, value, index) => sum + value * otherRow[index], 0)));
  const xty = transposed.map((row) => row.reduce((sum, value, index) => sum + value * y[index], 0));
  const inverse = webSheetFormulaMatrixInverse(xtx);
  if (inverse instanceof CellError) return inverse;
  const beta = webSheetFormulaMatrixVectorMultiply(inverse, xty).map(webSheetFormulaCleanMatrixNumber);
  const coefficients = beta.slice(0, xSpec.featureCount);
  const intercept = constant ? beta[beta.length - 1] : 0;
  const predicted = design.map((row) => row.reduce((sum, value, index) => sum + value * beta[index], 0));
  const residuals = y.map((value, index) => value - predicted[index]);
  const ssResidual = webSheetFormulaCleanMatrixNumber(residuals.reduce((sum, value) => sum + value * value, 0));
  const meanY = y.reduce((sum, value) => sum + value, 0) / y.length;
  const ssTotal = webSheetFormulaCleanMatrixNumber(
    constant ? y.reduce((sum, value) => sum + (value - meanY) ** 2, 0) : y.reduce((sum, value) => sum + value * value, 0),
  );
  const ssRegression = Math.max(0, webSheetFormulaCleanMatrixNumber(ssTotal - ssResidual));
  const df = observations.length - parameterCount;
  const mse = df > 0 ? ssResidual / df : NaN;
  const standardErrors = inverse.map((row, index) => {
    if (!(df > 0) || !Number.isFinite(mse)) return webSheetFormulaError(ErrorType.DIV_BY_ZERO);
    return webSheetFormulaCleanMatrixNumber(Math.sqrt(Math.max(0, mse * row[index])));
  });
  return {
    coefficients,
    intercept,
    beta,
    featureCount: xSpec.featureCount,
    constant,
    inverse,
    observations,
    outputShape: xSpec.outputShape,
    sourceHeight: xSpec.sourceHeight,
    sourceWidth: xSpec.sourceWidth,
    ssResidual,
    ssRegression,
    ssTotal,
    df,
    standardErrorY: df > 0 ? webSheetFormulaCleanMatrixNumber(Math.sqrt(mse)) : webSheetFormulaError(ErrorType.DIV_BY_ZERO),
    standardErrors,
    rSquared: ssTotal ? webSheetFormulaCleanMatrixNumber(1 - ssResidual / ssTotal) : 1,
    fStatistic: df > 0 && ssResidual ? webSheetFormulaCleanMatrixNumber((ssRegression / Math.max(1, xSpec.featureCount)) / (ssResidual / df)) : webSheetFormulaError(ErrorType.DIV_BY_ZERO),
  };
}



function webSheetFormulaRegressionCoefficientRow(fit, logOutput = false) {
  const slopes = fit.coefficients.slice().reverse();
  const values = fit.constant ? [...slopes, fit.intercept] : [...slopes, 0];
  return values.map((value) => webSheetFormulaCleanMatrixNumber(logOutput ? Math.exp(value) : value));
}



function webSheetFormulaRegressionStatsRow(fit, values) {
  const width = fit.featureCount + 1;
  return Array.from({ length: width }, (_unused, index) => index < values.length ? values[index] : "");
}



function webSheetFormulaRegressionResultMatrix(fit, options = {}) {
  const row1 = webSheetFormulaRegressionCoefficientRow(fit, Boolean(options.logOutput));
  if (!options.stats) return [row1];
  const stderr = fit.standardErrors.slice(0, fit.featureCount).reverse();
  const interceptStderr = fit.constant ? fit.standardErrors[fit.standardErrors.length - 1] : webSheetFormulaError(ErrorType.NA);
  return [
    row1,
    webSheetFormulaRegressionStatsRow(fit, [...stderr, interceptStderr]),
    webSheetFormulaRegressionStatsRow(fit, [fit.rSquared, fit.standardErrorY]),
    webSheetFormulaRegressionStatsRow(fit, [fit.fStatistic, fit.df]),
    webSheetFormulaRegressionStatsRow(fit, [fit.ssRegression, fit.ssResidual]),
  ];
}



function webSheetFormulaPredictionRowsFromMatrix(newX, fit) {
  if (newX === WEBSHEET_FORMULA_NA_SENTINEL || newX == null) return null;
  const matrix = webSheetFormulaRegressionNumericMatrix(newX);
  if (matrix instanceof CellError) return matrix;
  const { rows, height, width } = matrix;
  if (fit.featureCount === 1) {
    return {
      rows: rows.flatMap((row) => row).map((item) => [item]),
      shape: height === 1 && width > 1 ? "row" : "column",
      height,
      width,
    };
  }
  if (width === fit.featureCount) return { rows: rows.map((row) => row.slice()), shape: "column", height, width: 1 };
  if (height === fit.featureCount) {
    return {
      rows: Array.from({ length: width }, (_unused, colIndex) => rows.map((row) => row[colIndex])),
      shape: "row",
      height: 1,
      width,
    };
  }
  return webSheetFormulaError(ErrorType.VALUE);
}



function webSheetFormulaRegressionPredict(fit, newX = WEBSHEET_FORMULA_NA_SENTINEL, options = {}) {
  let predictionSpec = webSheetFormulaPredictionRowsFromMatrix(newX, fit);
  if (predictionSpec instanceof CellError) return predictionSpec;
  if (!predictionSpec) {
    predictionSpec = {
      rows: fit.observations.map((observation) => observation.x.slice()),
      shape: fit.outputShape,
      height: fit.sourceHeight,
      width: fit.sourceWidth,
    };
  }
  const predictions = [];
  for (const row of predictionSpec.rows) {
    if (row.length !== fit.featureCount || row.some((value) => !Number.isFinite(value))) return webSheetFormulaError(ErrorType.VALUE);
    const linear = fit.intercept + fit.coefficients.reduce((sum, coefficient, index) => sum + coefficient * row[index], 0);
    predictions.push(webSheetFormulaCleanMatrixNumber(options.exp ? Math.exp(linear) : linear));
  }
  if (predictionSpec.shape === "row") return [predictions];
  return predictions.map((value) => [value]);
}



function webSheetFormulaFrequency(dataArray, binsArray) {
  const data = webSheetFormulaNumericValues([dataArray]);
  if (data instanceof CellError) return data;
  const bins = webSheetFormulaNumericValues([binsArray]);
  if (bins instanceof CellError) return bins;
  const sortedBins = bins.filter(Number.isFinite).sort((left, right) => left - right);
  const counts = Array.from({ length: sortedBins.length + 1 }, () => 0);
  for (const value of data) {
    if (!Number.isFinite(value)) continue;
    const index = sortedBins.findIndex((bin) => value <= bin);
    counts[index >= 0 ? index : sortedBins.length] += 1;
  }
  return counts;
}



function webSheetFormulaRoundToDecimals(value, decimals = 0) {
  const places = Math.trunc(Number(decimals) || 0);
  if (places >= 0) {
    const factor = 10 ** places;
    return Math.round((Number(value) + Number.EPSILON) * factor) / factor;
  }
  const factor = 10 ** Math.abs(places);
  return Math.round(Number(value) / factor) * factor;
}



function webSheetFormulaFormatFixed(value, decimals = 2, useCommas = true) {
  const places = Math.max(0, Math.min(30, Math.trunc(Number(decimals) || 0)));
  const rounded = webSheetFormulaRoundToDecimals(value, decimals);
  const absolute = Math.abs(rounded);
  const text = absolute.toFixed(places);
  const formatted = useCommas ? text.replace(/\B(?=(\d{3})+(?!\d))/g, ",") : text;
  return rounded < 0 ? `-${formatted}` : formatted;
}



function webSheetFormulaMatrix(value) {
  if (value instanceof SimpleRangeValue) return value.rawData().map((row) => row.map((item) => webSheetFormulaRawValue(item)));
  if (Array.isArray(value)) {
    if (value.some(Array.isArray)) return value.map((row) => (Array.isArray(row) ? row : [row]).map((item) => webSheetFormulaRawValue(item)));
    return [value.map((item) => webSheetFormulaRawValue(item))];
  }
  return [[webSheetFormulaRawValue(value)]];
}



function webSheetFormulaMatrixDimensions(matrix) {
  return {
    height: matrix.length,
    width: Math.max(0, ...matrix.map((row) => row.length)),
  };
}



function webSheetFormulaNumberForMatrix(value) {
  const raw = webSheetFormulaRawValue(value);
  if (raw instanceof CellError) return raw;
  if (webSheetFormulaIsEmptyValue(raw)) return 0;
  const number = webSheetFormulaNumber(raw);
  return Number.isFinite(number) ? number : webSheetFormulaError(ErrorType.VALUE);
}



function webSheetFormulaNumericMatrix(value) {
  const matrix = webSheetFormulaMatrix(value);
  const { height, width } = webSheetFormulaMatrixDimensions(matrix);
  if (!height || !width) return webSheetFormulaError(ErrorType.VALUE);
  const numeric = [];
  for (let rowIndex = 0; rowIndex < height; rowIndex += 1) {
    const row = matrix[rowIndex] || [];
    if (row.length !== width) return webSheetFormulaError(ErrorType.VALUE);
    const numericRow = [];
    for (let colIndex = 0; colIndex < width; colIndex += 1) {
      const number = webSheetFormulaNumberForMatrix(row[colIndex]);
      if (number instanceof CellError) return number;
      numericRow.push(number);
    }
    numeric.push(numericRow);
  }
  return numeric;
}



function webSheetFormulaCleanMatrixNumber(value) {
  return Math.abs(value) < 1e-12 ? 0 : value;
}



function webSheetFormulaMatrixMultiply(leftValue, rightValue) {
  const left = webSheetFormulaNumericMatrix(leftValue);
  if (left instanceof CellError) return left;
  const right = webSheetFormulaNumericMatrix(rightValue);
  if (right instanceof CellError) return right;
  const leftDims = webSheetFormulaMatrixDimensions(left);
  const rightDims = webSheetFormulaMatrixDimensions(right);
  if (leftDims.width !== rightDims.height) return webSheetFormulaError(ErrorType.VALUE);
  const result = Array.from({ length: leftDims.height }, () => Array(rightDims.width).fill(0));
  for (let row = 0; row < leftDims.height; row += 1) {
    for (let col = 0; col < rightDims.width; col += 1) {
      let sum = 0;
      for (let index = 0; index < leftDims.width; index += 1) {
        sum += left[row][index] * right[index][col];
      }
      result[row][col] = webSheetFormulaCleanMatrixNumber(sum);
    }
  }
  return result;
}



function webSheetFormulaMatrixDeterminant(value) {
  const matrix = webSheetFormulaNumericMatrix(value);
  if (matrix instanceof CellError) return matrix;
  const { height, width } = webSheetFormulaMatrixDimensions(matrix);
  if (height !== width) return webSheetFormulaError(ErrorType.VALUE);
  const work = matrix.map((row) => row.slice());
  let sign = 1;
  let det = 1;
  for (let pivot = 0; pivot < height; pivot += 1) {
    let pivotRow = pivot;
    let pivotAbs = Math.abs(work[pivot][pivot]);
    for (let row = pivot + 1; row < height; row += 1) {
      const candidateAbs = Math.abs(work[row][pivot]);
      if (candidateAbs > pivotAbs) {
        pivotAbs = candidateAbs;
        pivotRow = row;
      }
    }
    if (pivotAbs < 1e-12) return 0;
    if (pivotRow !== pivot) {
      [work[pivot], work[pivotRow]] = [work[pivotRow], work[pivot]];
      sign *= -1;
    }
    const pivotValue = work[pivot][pivot];
    det *= pivotValue;
    for (let row = pivot + 1; row < height; row += 1) {
      const factor = work[row][pivot] / pivotValue;
      for (let col = pivot + 1; col < height; col += 1) {
        work[row][col] -= factor * work[pivot][col];
      }
    }
  }
  return webSheetFormulaCleanMatrixNumber(det * sign);
}



function webSheetFormulaMatrixInverse(value) {
  const matrix = webSheetFormulaNumericMatrix(value);
  if (matrix instanceof CellError) return matrix;
  const { height, width } = webSheetFormulaMatrixDimensions(matrix);
  if (height !== width) return webSheetFormulaError(ErrorType.VALUE);
  const size = height;
  const work = matrix.map((row, rowIndex) => [
    ...row,
    ...Array.from({ length: size }, (_unused, colIndex) => (rowIndex === colIndex ? 1 : 0)),
  ]);
  for (let pivot = 0; pivot < size; pivot += 1) {
    let pivotRow = pivot;
    let pivotAbs = Math.abs(work[pivot][pivot]);
    for (let row = pivot + 1; row < size; row += 1) {
      const candidateAbs = Math.abs(work[row][pivot]);
      if (candidateAbs > pivotAbs) {
        pivotAbs = candidateAbs;
        pivotRow = row;
      }
    }
    if (pivotAbs < 1e-12) return webSheetFormulaError(ErrorType.NUM);
    if (pivotRow !== pivot) [work[pivot], work[pivotRow]] = [work[pivotRow], work[pivot]];
    const pivotValue = work[pivot][pivot];
    for (let col = 0; col < size * 2; col += 1) work[pivot][col] /= pivotValue;
    for (let row = 0; row < size; row += 1) {
      if (row === pivot) continue;
      const factor = work[row][pivot];
      for (let col = 0; col < size * 2; col += 1) {
        work[row][col] -= factor * work[pivot][col];
      }
    }
  }
  return work.map((row) => row.slice(size).map(webSheetFormulaCleanMatrixNumber));
}



function webSheetFormulaMatrixIdentity(sizeValue) {
  const size = Math.trunc(Number(webSheetFormulaRawValue(sizeValue)));
  if (!Number.isFinite(size) || size < 1) return webSheetFormulaError(ErrorType.VALUE);
  return Array.from({ length: size }, (_row, rowIndex) =>
    Array.from({ length: size }, (_col, colIndex) => (rowIndex === colIndex ? 1 : 0)),
  );
}



function webSheetFormulaNormalizeKey(value) {
  const raw = webSheetFormulaRawValue(value);
  if (raw instanceof CellError) return `error:${raw.type}`;
  if (webSheetFormulaIsEmptyValue(raw)) return "empty:";
  if (typeof raw === "number") return `number:${Object.is(raw, -0) ? 0 : raw}`;
  if (typeof raw === "boolean") return `boolean:${raw}`;
  return `string:${String(raw).toLocaleLowerCase()}`;
}



function webSheetFormulaRowKey(row) {
  return row.map(webSheetFormulaNormalizeKey).join("\u0001");
}



function webSheetFormulaArraySizeFromFirstArg(ast, state, plugin) {
  if (!ast.args.length) return ArraySize.error();
  const size = plugin.arraySizeForAst(ast.args[0], state);
  return size?.isRef ? new ArraySize(Math.max(1, size.width), Math.max(1, size.height)) : size;
}



function webSheetFormulaAstNumberLiteral(ast, fallback = 1) {
  if (!ast) return fallback;
  if (ast.type === "NUMBER") return Number(ast.value);
  if (ast.type === "PLUS_UNARY_OP") return webSheetFormulaAstNumberLiteral(ast.value, fallback);
  if (ast.type === "MINUS_UNARY_OP") return -webSheetFormulaAstNumberLiteral(ast.value, fallback);
  if (ast.type === "PARENTHESES") return webSheetFormulaAstNumberLiteral(ast.expression, fallback);
  return fallback;
}



function webSheetFormulaAstStringLiteral(ast) {
  if (!ast) return "";
  if (ast.type === "STRING") return String(ast.value ?? "");
  if (ast.type === "PARENTHESES") return webSheetFormulaAstStringLiteral(ast.expression);
  return "";
}



function webSheetFormulaAstStaticValue(ast, fallback = undefined) {
  if (!ast || ast.type === "EMPTY") return fallback;
  if (ast.type === "STRING") return String(ast.value ?? "");
  if (ast.type === "NUMBER") return Number(ast.value);
  if (ast.type === "BOOLEAN") return Boolean(ast.value);
  if (ast.type === "NAMED_EXPRESSION" && /^(TRUE|FALSE)$/iu.test(ast.expressionName || "")) return /^TRUE$/iu.test(ast.expressionName);
  if (ast.type === "PLUS_UNARY_OP") return webSheetFormulaAstStaticValue(ast.value, fallback);
  if (ast.type === "MINUS_UNARY_OP") {
    const value = webSheetFormulaAstStaticValue(ast.value, fallback);
    return Number.isFinite(Number(value)) ? -Number(value) : fallback;
  }
  if (ast.type === "PARENTHESES") return webSheetFormulaAstStaticValue(ast.expression, fallback);
  return fallback;
}



function webSheetFormulaAstStaticMatrix(ast) {
  if (!ast || ast.type === "EMPTY") return null;
  if (ast.type === "ARRAY") {
    return ast.args.map((row) => row.map((cellAst) => webSheetFormulaAstStaticValue(cellAst, "")));
  }
  const scalar = webSheetFormulaAstStaticValue(ast, undefined);
  return scalar === undefined ? null : [[scalar]];
}



function webSheetFormulaAstRuntimeScalar(ast, state, plugin, fallback = undefined) {
  if (!ast || ast.type === "EMPTY") return fallback;
  try {
    const value = plugin.evaluateAst(ast, state);
    if (value instanceof SimpleRangeValue) return value.rawData()?.[0]?.[0] ?? fallback;
    return webSheetFormulaRawValue(value);
  } catch {
    return fallback;
  }
}



function webSheetFormulaAstRuntimeMatrix(ast, state, plugin, fallback = [[]]) {
  if (!ast || ast.type === "EMPTY") return fallback;
  if (ast.type === "ARRAY") {
    return ast.args.map((row) => row.map((cellAst) => webSheetFormulaAstRuntimeScalar(cellAst, state, plugin, "")));
  }
  const value = plugin.evaluateAst(ast, state);
  if (value instanceof SimpleRangeValue) return value.rawData().map((row) => row.map((item) => webSheetFormulaRawValue(item)));
  return [[webSheetFormulaRawValue(value)]];
}



function webSheetFormulaAstBooleanLiteral(ast, fallback = false) {
  if (!ast) return fallback;
  if (ast.type === "BOOLEAN") return Boolean(ast.value);
  if (ast.type === "NUMBER") return Number(ast.value) !== 0;
  if (ast.type === "NAMED_EXPRESSION" && /^(TRUE|FALSE)$/iu.test(ast.expressionName || "")) return /^TRUE$/iu.test(ast.expressionName);
  if (ast.type === "PARENTHESES") return webSheetFormulaAstBooleanLiteral(ast.expression, fallback);
  return fallback;
}



function webSheetFormulaRegressionFeatureCountFromSizes(ySize, xSize) {
  if (!xSize) return 1;
  const yLength = Math.max(1, (ySize?.width || 1) * (ySize?.height || 1));
  if (xSize.height === yLength) return Math.max(1, xSize.width);
  if (xSize.width === yLength) return Math.max(1, xSize.height);
  return 1;
}



function webSheetFormulaCompareForSort(left, right) {
  const leftEmpty = webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(left));
  const rightEmpty = webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(right));
  if (leftEmpty && rightEmpty) return 0;
  if (leftEmpty) return 1;
  if (rightEmpty) return -1;
  const comparison = webSheetFormulaCompareValues(left, right);
  return Number.isFinite(comparison) ? comparison : 0;
}



function webSheetFormulaLookupEntries(lookupRange, resultRange) {
  const lookupData = webSheetFormulaMatrix(lookupRange);
  const resultData = resultRange == null ? null : webSheetFormulaMatrix(resultRange);
  const { height, width } = webSheetFormulaMatrixDimensions(lookupData);
  if (!height || !width) return [];
  if (resultData) {
    const lookupValues = lookupData.flat();
    const resultValues = resultData.flat();
    return lookupValues.slice(0, resultValues.length).map((lookup, index) => ({ lookup, result: resultValues[index] }));
  }
  if (height >= width) {
    return lookupData.map((row) => ({ lookup: row[0], result: row[Math.max(0, width - 1)] }));
  }
  const firstRow = lookupData[0] || [];
  const lastRow = lookupData[Math.max(0, height - 1)] || [];
  return firstRow.map((lookup, index) => ({ lookup, result: lastRow[index] }));
}



function webSheetFormulaApproximateLookup(lookupValue, entries) {
  const target = webSheetFormulaRawValue(lookupValue);
  let best = null;
  for (const entry of entries) {
    if (entry.lookup instanceof CellError || webSheetFormulaIsEmptyValue(entry.lookup)) continue;
    const comparison = webSheetFormulaCompareValues(entry.lookup, target);
    if (!Number.isFinite(comparison) || comparison > 0) continue;
    if (!best || webSheetFormulaCompareValues(entry.lookup, best.lookup) >= 0) best = entry;
  }
  return best ? best.result : webSheetFormulaError(ErrorType.NA);
}



function webSheetFormulaLookupMatchIndex(lookupValue, lookupArray, matchMode = 0, searchMode = 1) {
  const normalizedMatchMode = Math.trunc(Number(matchMode ?? 0));
  const normalizedSearchMode = Math.trunc(Number(searchMode ?? 1));
  if (![0, -1, 1, 2].includes(normalizedMatchMode) || ![1, -1, 2, -2].includes(normalizedSearchMode)) {
    return { error: webSheetFormulaError(ErrorType.VALUE), index: -1 };
  }
  const values = webSheetFormulaFlattenValues(lookupArray).map(webSheetFormulaRawValue);
  if (Math.abs(normalizedSearchMode) === 2 && normalizedMatchMode !== 2) {
    return { index: webSheetFormulaBinaryMatchIndex(lookupValue, values, normalizedMatchMode, normalizedSearchMode < 0) };
  }
  const reverse = normalizedSearchMode === -1 || normalizedSearchMode === -2;
  const exactIndex = webSheetFormulaExactMatchIndex(lookupValue, values, normalizedMatchMode === 2, reverse);
  if (exactIndex >= 0) return { index: exactIndex };
  if (normalizedMatchMode === 0 || normalizedMatchMode === 2) return { index: -1 };
  return { index: webSheetFormulaNearestMatchIndex(lookupValue, values, normalizedMatchMode, reverse) };
}



function webSheetFormulaXlookupMatrix(lookupValue, lookupArray, returnArray, ifNotFound, matchMode = 0, searchMode = 1) {
  const match = webSheetFormulaLookupMatchIndex(lookupValue, lookupArray, matchMode, searchMode);
  if (match.error) return [[match.error]];
  if (match.index < 0) {
    return [[ifNotFound === WEBSHEET_FORMULA_NA_SENTINEL ? webSheetFormulaError(ErrorType.NA) : ifNotFound]];
  }

  const lookupData = webSheetFormulaMatrix(lookupArray);
  const returnData = webSheetFormulaMatrix(returnArray);
  const lookupDims = webSheetFormulaMatrixDimensions(lookupData);
  const returnDims = webSheetFormulaMatrixDimensions(returnData);
  if (!returnDims.height || !returnDims.width) return [[webSheetFormulaError(ErrorType.VALUE)]];

  if (lookupDims.height === 1 && lookupDims.width > 1) {
    if (returnDims.height === 1) return [[returnData[0]?.[match.index] ?? webSheetFormulaError(ErrorType.NA)]];
    return returnData.map((row) => [row[match.index] ?? ""]);
  }

  if (returnDims.height > match.index) {
    const row = returnData[match.index] || [];
    return [row.length ? row : [""]];
  }

  const flattened = webSheetFormulaFlattenValues(returnArray).map(webSheetFormulaRawValue);
  return [[flattened[match.index] ?? webSheetFormulaError(ErrorType.NA)]];
}



function webSheetFormulaXlookupResult(lookupValue, lookupArray, returnArray, ifNotFound, matchMode = 0, searchMode = 1) {
  const matrix = webSheetFormulaXlookupMatrix(lookupValue, lookupArray, returnArray, ifNotFound, matchMode, searchMode);
  const { height, width } = webSheetFormulaMatrixDimensions(matrix);
  return height === 1 && width === 1 ? matrix[0]?.[0] ?? "" : SimpleRangeValue.onlyValues(matrix);
}



function webSheetFormulaTextDelimiters(delimiterValue) {
  const delimiters = webSheetFormulaFlattenValues(delimiterValue)
    .map((value) => webSheetFormulaText(value))
    .filter((text) => !(text instanceof CellError) && text !== "");
  return delimiters.length ? delimiters : [""];
}



function webSheetFormulaSplitTextByDelimiters(text, delimiters, matchMode = 0) {
  const source = String(text ?? "");
  const items = webSheetFormulaTextDelimiters(delimiters);
  if (!items.length || items.some((item) => item === "")) return [source];
  const flags = matchMode ? "iu" : "u";
  const pattern = items
    .sort((left, right) => right.length - left.length)
    .map((item) => item.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join("|");
  return source.split(new RegExp(pattern, flags));
}



function textSplitFormulaMatrix(text, colDelimiter, rowDelimiter = undefined, ignoreEmpty = false, matchMode = 0, padWith = webSheetFormulaError(ErrorType.NA)) {
  const source = webSheetFormulaText(text);
  if (source instanceof CellError) return [[source]];
  const padValue = padWith === WEBSHEET_FORMULA_NA_SENTINEL || padWith == null ? webSheetFormulaError(ErrorType.NA) : padWith;
  const rowParts = rowDelimiter == null || webSheetFormulaTextDelimiters(rowDelimiter).every((item) => item === "")
    ? [source]
    : webSheetFormulaSplitTextByDelimiters(source, rowDelimiter, matchMode);
  const rows = rowParts.map((rowText) => {
    let cells = webSheetFormulaSplitTextByDelimiters(rowText, colDelimiter, matchMode);
    if (ignoreEmpty) cells = cells.filter((item) => item !== "");
    return cells;
  });
  const width = Math.max(1, ...rows.map((row) => row.length));
  return rows.map((row) => {
    const normalized = row.length ? [...row] : [""];
    while (normalized.length < width) normalized.push(padValue);
    return normalized;
  });
}



function textPartFormulaValue(text, delimiter, instanceNum = 1, matchMode = 0, matchEnd = false, ifNotFound = webSheetFormulaError(ErrorType.NA), after = false) {
  const source = webSheetFormulaText(text);
  if (source instanceof CellError) return source;
  const instance = Math.trunc(Number(instanceNum == null || String(instanceNum).trim() === "" ? 1 : instanceNum));
  if (!instance) return webSheetFormulaError(ErrorType.VALUE);
  const mode = Math.trunc(Number(matchMode == null || String(matchMode).trim() === "" ? 0 : matchMode));
  const fallback = ifNotFound === WEBSHEET_FORMULA_NA_SENTINEL || ifNotFound == null ? webSheetFormulaError(ErrorType.NA) : ifNotFound;
  const delimiters = webSheetFormulaTextDelimiters(delimiter);
  if (!delimiters.length || delimiters.some((item) => item === "")) return webSheetFormulaError(ErrorType.VALUE);
  const sourceForSearch = mode ? source.toLocaleLowerCase("ja-JP") : source;
  const normalizedDelimiters = delimiters.map((item) => (mode ? item.toLocaleLowerCase("ja-JP") : item));
  const matches = [];
  for (let index = 0; index < sourceForSearch.length; index += 1) {
    const delimiterIndex = normalizedDelimiters.findIndex((item) => sourceForSearch.startsWith(item, index));
    if (delimiterIndex < 0) continue;
    matches.push({ index, length: delimiters[delimiterIndex].length });
    index += Math.max(0, delimiters[delimiterIndex].length - 1);
  }
  if (formulaArgumentBoolean(matchEnd, false)) matches.push({ index: source.length, length: 0 });
  const match = instance > 0 ? matches[instance - 1] : matches[matches.length + instance];
  if (!match) return fallback;
  return after ? source.slice(match.index + match.length) : source.slice(0, match.index);
}



function webSheetFormulaExactMatchIndex(lookupValue, values, wildcard = false, reverse = false) {
  const source = reverse ? [...values].reverse() : values;
  const offset = reverse ? values.length - 1 : 0;
  const step = reverse ? -1 : 1;
  const matcher =
    wildcard && typeof webSheetFormulaRawValue(lookupValue) === "string"
      ? webSheetFormulaWildcardRegex(String(webSheetFormulaRawValue(lookupValue)).replace(/~([*?])/gu, "$1"))
      : null;
  for (let index = 0; index < source.length; index += 1) {
    const value = source[index];
    if (value instanceof CellError) continue;
    if (matcher) {
      if (matcher.test(webSheetFormulaText(value))) return offset + index * step;
      continue;
    }
    if (webSheetFormulaCompareValues(value, lookupValue) === 0) return offset + index * step;
  }
  return -1;
}



function webSheetFormulaNearestMatchIndex(lookupValue, values, matchMode, reverse = false) {
  let bestIndex = -1;
  const source = reverse ? [...values].reverse() : values;
  const offset = reverse ? values.length - 1 : 0;
  const step = reverse ? -1 : 1;
  for (let index = 0; index < source.length; index += 1) {
    const value = source[index];
    const originalIndex = offset + index * step;
    if (value instanceof CellError || webSheetFormulaIsEmptyValue(value)) continue;
    const comparison = webSheetFormulaCompareValues(value, lookupValue);
    if (!Number.isFinite(comparison)) continue;
    if (comparison === 0) return originalIndex;
    if (matchMode < 0 && comparison < 0) {
      if (bestIndex < 0 || webSheetFormulaCompareValues(value, values[bestIndex]) > 0) bestIndex = originalIndex;
    } else if (matchMode > 0 && comparison > 0) {
      if (bestIndex < 0 || webSheetFormulaCompareValues(value, values[bestIndex]) < 0) bestIndex = originalIndex;
    }
  }
  return bestIndex;
}



function webSheetFormulaBinaryMatchIndex(lookupValue, values, matchMode = 0, descending = false) {
  let low = 0;
  let high = values.length - 1;
  let exact = -1;
  let nearest = -1;
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    const value = values[mid];
    if (value instanceof CellError || webSheetFormulaIsEmptyValue(value)) {
      return webSheetFormulaNearestMatchIndex(lookupValue, values, matchMode, descending);
    }
    const comparison = webSheetFormulaCompareValues(value, lookupValue);
    if (!Number.isFinite(comparison)) {
      return webSheetFormulaNearestMatchIndex(lookupValue, values, matchMode, descending);
    }
    if (comparison === 0) {
      exact = mid;
      high = mid - 1;
      continue;
    }
    if (!descending) {
      if (matchMode < 0 && comparison < 0) {
        nearest = mid;
        low = mid + 1;
      } else if (matchMode > 0 && comparison > 0) {
        nearest = mid;
        high = mid - 1;
      } else if (comparison < 0) {
        low = mid + 1;
      } else {
        high = mid - 1;
      }
    } else if (matchMode < 0 && comparison < 0) {
      nearest = mid;
      high = mid - 1;
    } else if (matchMode > 0 && comparison > 0) {
      nearest = mid;
      low = mid + 1;
    } else if (comparison > 0) {
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  if (exact >= 0) return exact;
  return matchMode === 0 ? -1 : nearest;
}



function webSheetFormulaConvertUnitKey(unit) {
  const text = String(unit ?? "").trim();
  if (text === "C" || /^cel(?:sius)?$/iu.test(text)) return "c";
  if (text === "F" || /^fah(?:renheit)?$/iu.test(text)) return "f";
  if (text === "K" || /^kelvin$/iu.test(text)) return "k";
  return text.toLocaleLowerCase();
}



function webSheetFormulaConvertTemperature(number, fromUnit, toUnit) {
  const toKelvin = {
    c: (value) => value + 273.15,
    f: (value) => (value - 32) * (5 / 9) + 273.15,
    k: (value) => value,
  };
  const fromKelvin = {
    c: (value) => value - 273.15,
    f: (value) => (value - 273.15) * (9 / 5) + 32,
    k: (value) => value,
  };
  if (!toKelvin[fromUnit] || !fromKelvin[toUnit]) return null;
  return fromKelvin[toUnit](toKelvin[fromUnit](number));
}



function webSheetFormulaDatabaseFieldIndex(database, field) {
  const headers = database[0] || [];
  const numericField = webSheetFormulaNumber(field);
  if (Number.isFinite(numericField)) {
    const index = Math.trunc(numericField) - 1;
    return index >= 0 && index < headers.length ? index : -1;
  }
  const fieldText = webSheetFormulaText(field);
  if (fieldText instanceof CellError) return -1;
  const normalized = String(fieldText).trim().toLocaleLowerCase();
  return headers.findIndex((header) => String(webSheetFormulaRawValue(header) ?? "").trim().toLocaleLowerCase() === normalized);
}



function webSheetFormulaDatabaseRowMatches(row, headers, criteria) {
  if (!criteria.length || criteria.length < 2) return true;
  const criteriaHeaders = criteria[0] || [];
  const criteriaRows = criteria.slice(1).filter((criteriaRow) => criteriaRow.some((value) => !webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(value))));
  if (!criteriaRows.length) return true;
  return criteriaRows.some((criteriaRow) =>
    criteriaRow.every((criterion, colIndex) => {
      if (webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(criterion))) return true;
      const headerText = String(webSheetFormulaRawValue(criteriaHeaders[colIndex]) ?? "").trim().toLocaleLowerCase();
      const fieldIndex = headers.findIndex((header) => String(webSheetFormulaRawValue(header) ?? "").trim().toLocaleLowerCase() === headerText);
      if (fieldIndex < 0) return false;
      return webSheetFormulaCriteriaMatcher(criterion)(row[fieldIndex]);
    }),
  );
}



function webSheetFormulaDatabaseValues(databaseRange, field, criteriaRange) {
  const database = webSheetFormulaMatrix(databaseRange);
  const criteria = webSheetFormulaMatrix(criteriaRange);
  if (database.length < 2) return webSheetFormulaError(ErrorType.VALUE);
  const headers = database[0] || [];
  const fieldIndex = webSheetFormulaDatabaseFieldIndex(database, field);
  if (fieldIndex < 0) return webSheetFormulaError(ErrorType.VALUE);
  const values = [];
  for (const row of database.slice(1)) {
    if (webSheetFormulaDatabaseRowMatches(row, headers, criteria)) values.push(row[fieldIndex]);
  }
  return values;
}



function webSheetFormulaDatabaseAggregate(databaseRange, field, criteriaRange, mode) {
  const values = webSheetFormulaDatabaseValues(databaseRange, field, criteriaRange);
  if (values instanceof CellError) return values;
  if (mode === "counta") return values.filter((value) => !webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(value))).length;
  const numbers = [];
  for (const value of values) {
    const number = webSheetFormulaNumber(value);
    if (number instanceof CellError) return number;
    if (Number.isFinite(number)) numbers.push(number);
  }
  if (mode === "count") return numbers.length;
  if (mode === "sum") return numbers.reduce((sum, value) => sum + value, 0);
  if (mode === "average") return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : webSheetFormulaError(ErrorType.DIV_BY_ZERO);
  if (mode === "max") return numbers.length ? Math.max(...numbers) : 0;
  if (mode === "min") return numbers.length ? Math.min(...numbers) : 0;
  if (mode === "product") return numbers.length ? numbers.reduce((product, value) => product * value, 1) : 0;
  if (mode === "var" || mode === "stdev" || mode === "varp" || mode === "stdevp") {
    const sample = mode === "var" || mode === "stdev";
    const variance = webSheetFormulaVariance(numbers, sample);
    if (variance instanceof CellError) return variance;
    return mode === "stdev" || mode === "stdevp" ? Math.sqrt(variance) : variance;
  }
  return webSheetFormulaError(ErrorType.VALUE);
}



function webSheetFormulaUnwrapAstParenthesis(arg) {
  let current = arg;
  while (current?.type === "PARENTHESES" || current?.type === "PARENTHESIS") current = current.expression;
  return current;
}



function webSheetFormulaSheetIndexFromFormulaSheetId(sheetId) {
  const sheetName = state.hf?.getSheetName?.(sheetId);
  const byName = sheetName ? state.model?.sheets?.findIndex((sheet) => sheet.name === sheetName) : -1;
  if (byName >= 0) return byName;
  const numeric = Number(sheetId);
  return Number.isInteger(numeric) && numeric >= 0 && numeric < (state.model?.sheets?.length || 0) ? numeric : state.activeSheetIndex;
}



function webSheetFormulaReferenceFromAstArg(arg, formulaAddress = {}) {
  const unwrapped = webSheetFormulaUnwrapAstParenthesis(arg);
  if (!unwrapped) return null;
  try {
    if (unwrapped.type === "CELL_REFERENCE") {
      const ref = unwrapped.reference?.toSimpleCellAddress?.(formulaAddress);
      if (!ref) return null;
      return {
        kind: "cell",
        sheetIndex: webSheetFormulaSheetIndexFromFormulaSheetId(ref.sheet),
        top: Number(ref.row) + 1,
        left: Number(ref.col) + 1,
        bottom: Number(ref.row) + 1,
        right: Number(ref.col) + 1,
      };
    }
    if (unwrapped.type === "CELL_RANGE") {
      const start = unwrapped.start?.toSimpleCellAddress?.(formulaAddress);
      const end = unwrapped.end?.toSimpleCellAddress?.(formulaAddress);
      if (!start || !end) return null;
      return {
        kind: "range",
        sheetIndex: webSheetFormulaSheetIndexFromFormulaSheetId(start.sheet),
        top: Math.min(Number(start.row), Number(end.row)) + 1,
        left: Math.min(Number(start.col), Number(end.col)) + 1,
        bottom: Math.max(Number(start.row), Number(end.row)) + 1,
        right: Math.max(Number(start.col), Number(end.col)) + 1,
      };
    }
    if (unwrapped.type === "COLUMN_RANGE") {
      const start = unwrapped.start?.toSimpleColumnAddress?.(formulaAddress);
      const end = unwrapped.end?.toSimpleColumnAddress?.(formulaAddress);
      if (!start || !end) return null;
      const sheetIndex = webSheetFormulaSheetIndexFromFormulaSheetId(start.sheet);
      const sheet = state.model?.sheets?.[sheetIndex];
      return {
        kind: "columnRange",
        sheetIndex,
        top: 1,
        left: Math.min(Number(start.col), Number(end.col)) + 1,
        bottom: Math.max(1, sheet?.rows?.length || 1),
        right: Math.max(Number(start.col), Number(end.col)) + 1,
      };
    }
    if (unwrapped.type === "ROW_RANGE") {
      const start = unwrapped.start?.toSimpleRowAddress?.(formulaAddress);
      const end = unwrapped.end?.toSimpleRowAddress?.(formulaAddress);
      if (!start || !end) return null;
      const sheetIndex = webSheetFormulaSheetIndexFromFormulaSheetId(start.sheet);
      const sheet = state.model?.sheets?.[sheetIndex];
      return {
        kind: "rowRange",
        sheetIndex,
        top: Math.min(Number(start.row), Number(end.row)) + 1,
        left: 1,
        bottom: Math.max(Number(start.row), Number(end.row)) + 1,
        right: Math.max(1, sheet?.columns?.length || 1),
      };
    }
  } catch {
    return null;
  }
  return null;
}



function webSheetFormulaReferenceCell(range) {
  if (!range) return null;
  const sheet = state.model?.sheets?.[range.sheetIndex];
  if (!sheet) return null;
  return { sheet, sheetIndex: range.sheetIndex, row: range.top, col: range.left, cell: sheet.cells?.[cellKey(range.top, range.left)] || null };
}



function webSheetFormulaCellInfoType(cell) {
  if (cellHasFormulaContent(cell)) return "v";
  if (!cell || webSheetFormulaIsEmptyValue(cell.raw ?? cell.cached ?? cell.display)) return "b";
  return typeof cell.raw === "string" && !cell.raw.startsWith("=") ? "l" : "v";
}



function webSheetFormulaCellInfoFormat(cell) {
  const fmt = String(cell?.numFmt || "").toLocaleLowerCase();
  if (!fmt || fmt === "general") return "G";
  if (/[ymdhsa]/i.test(fmt)) return "D";
  if (/%/.test(fmt)) return "P";
  if (currencySymbol(fmt)) return "C";
  if (/0[.,]0|0\.0/.test(fmt)) return "F";
  return "G";
}



function webSheetFormulaWorkbookFileName() {
  const title = typeof document === "undefined" ? "" : document.title;
  const name = state.model?.fileName || state.currentFileName || state.sourceFileName || title || "";
  return name ? `[${name}]` : "";
}

export {
  daysInYearUtc,
  loadWorkbookIoModules,
  textPartFormulaValue,
  textSplitFormulaMatrix,
  warmWorkbookIoModules,
  webSheetFormulaAValues,
  webSheetFormulaAccrint,
  webSheetFormulaAccrintM,
  webSheetFormulaAcot,
  webSheetFormulaAddMonths,
  webSheetFormulaAggregate,
  webSheetFormulaAggregateValues,
  webSheetFormulaAmorDegrc,
  webSheetFormulaAmorDegrcCoefficient,
  webSheetFormulaAmorFirstFraction,
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
  webSheetFormulaAstStaticValue,
  webSheetFormulaAstStringLiteral,
  webSheetFormulaAverageDeviation,
  webSheetFormulaAverageIf,
  webSheetFormulaBase,
  webSheetFormulaBesselI,
  webSheetFormulaBesselI0,
  webSheetFormulaBesselI1,
  webSheetFormulaBesselJ,
  webSheetFormulaBesselJ0,
  webSheetFormulaBesselJ1,
  webSheetFormulaBesselK,
  webSheetFormulaBesselK0,
  webSheetFormulaBesselK1,
  webSheetFormulaBesselOrder,
  webSheetFormulaBesselY,
  webSheetFormulaBesselY0,
  webSheetFormulaBesselY1,
  webSheetFormulaBetaDist,
  webSheetFormulaBetaInv,
  webSheetFormulaBinaryMatchIndex,
  webSheetFormulaBinomDist,
  webSheetFormulaBinomDistRange,
  webSheetFormulaBinomInv,
  webSheetFormulaBitNumber,
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
  webSheetFormulaCleanMatrixNumber,
  webSheetFormulaCombination,
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
  webSheetFormulaCouponInfo,
  webSheetFormulaCouponPartialDays,
  webSheetFormulaCouponPeriodDays,
  webSheetFormulaCovariance,
  webSheetFormulaCriteriaMatcher,
  webSheetFormulaCriteriaPairs,
  webSheetFormulaDatabaseAggregate,
  webSheetFormulaDatabaseFieldIndex,
  webSheetFormulaDatabaseRowMatches,
  webSheetFormulaDatabaseValues,
  webSheetFormulaDateParts,
  webSheetFormulaDatePartsToSerial,
  webSheetFormulaDateSerial,
  webSheetFormulaDateSerialWeekday,
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
  webSheetFormulaDollarFractionDenominator,
  webSheetFormulaDoubleFactorial,
  webSheetFormulaErf,
  webSheetFormulaError,
  webSheetFormulaExactMatchIndex,
  webSheetFormulaExponDist,
  webSheetFormulaFDist,
  webSheetFormulaFInv,
  webSheetFormulaFTest,
  webSheetFormulaFactorial,
  webSheetFormulaFactorialResult,
  webSheetFormulaFinancialType,
  webSheetFormulaFinancialValues,
  webSheetFormulaFlattenValues,
  webSheetFormulaFloorMath,
  webSheetFormulaFloorPrecise,
  webSheetFormulaFormatFixed,
  webSheetFormulaFrequency,
  webSheetFormulaFutureValue,
  webSheetFormulaFvSchedule,
  webSheetFormulaGamma,
  webSheetFormulaGammaContinuedFractionQ,
  webSheetFormulaGammaDist,
  webSheetFormulaGammaInv,
  webSheetFormulaGammaSeriesP,
  webSheetFormulaGeoMean,
  webSheetFormulaGreatestCommonDivisor,
  webSheetFormulaHarMean,
  webSheetFormulaHolidaySet,
  webSheetFormulaHypergeomDist,
  webSheetFormulaImabs,
  webSheetFormulaImproduct,
  webSheetFormulaImsum,
  webSheetFormulaInfoN,
  webSheetFormulaIntrate,
  webSheetFormulaInverseByBisection,
  webSheetFormulaInverseStandardNormal,
  webSheetFormulaIpmt,
  webSheetFormulaIrr,
  webSheetFormulaIsEmptyValue,
  webSheetFormulaIsWorkday,
  webSheetFormulaIsoWeeknumValue,
  webSheetFormulaIspmt,
  webSheetFormulaKnownXRows,
  webSheetFormulaKthValue,
  webSheetFormulaKurt,
  webSheetFormulaLambdaSpec,
  webSheetFormulaLeastCommonMultiple,
  webSheetFormulaLetApplyBinary,
  webSheetFormulaLetAstContainsEnvName,
  webSheetFormulaLetAstValue,
  webSheetFormulaLetFunctionValue,
  webSheetFormulaLetLogical,
  webSheetFormulaLetName,
  webSheetFormulaLetScalarNumber,
  webSheetFormulaLogCombination,
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
  webSheetFormulaMatrixValues,
  webSheetFormulaMatrixVectorMultiply,
  webSheetFormulaMedian,
  webSheetFormulaMirr,
  webSheetFormulaMod,
  webSheetFormulaModeMultValues,
  webSheetFormulaModeSnglValue,
  webSheetFormulaMround,
  webSheetFormulaMultipleRegression,
  webSheetFormulaNearestMatchIndex,
  webSheetFormulaNegbinomDist,
  webSheetFormulaNetworkDays,
  webSheetFormulaNormalCdf,
  webSheetFormulaNormalInv,
  webSheetFormulaNormalPdf,
  webSheetFormulaNormalizeKey,
  webSheetFormulaNper,
  webSheetFormulaNpv,
  webSheetFormulaNumber,
  webSheetFormulaNumberForMatrix,
  webSheetFormulaNumericMatrix,
  webSheetFormulaNumericOrZero,
  webSheetFormulaNumericValues,
  webSheetFormulaPairedNumericValues,
  webSheetFormulaParseDateText,
  webSheetFormulaParseSignedBase,
  webSheetFormulaParseTimeText,
  webSheetFormulaPayment,
  webSheetFormulaPduration,
  webSheetFormulaPercentRank,
  webSheetFormulaPercentile,
  webSheetFormulaPermutation,
  webSheetFormulaPermutationA,
  webSheetFormulaPlacesValue,
  webSheetFormulaPoissonDist,
  webSheetFormulaPpmt,
  webSheetFormulaPredictionRowsFromMatrix,
  webSheetFormulaPresentValue,
  webSheetFormulaPriceDisc,
  webSheetFormulaPriceMat,
  webSheetFormulaProb,
  webSheetFormulaQuartile,
  webSheetFormulaRangeData,
  webSheetFormulaRank,
  webSheetFormulaRate,
  webSheetFormulaRawValue,
  webSheetFormulaReceived,
  webSheetFormulaReferenceCell,
  webSheetFormulaReferenceFromAstArg,
  webSheetFormulaRegression,
  webSheetFormulaRegressionCoefficientRow,
  webSheetFormulaRegressionFeatureCountFromSizes,
  webSheetFormulaRegressionNumericMatrix,
  webSheetFormulaRegressionNumericVector,
  webSheetFormulaRegressionPredict,
  webSheetFormulaRegressionRawMatrix,
  webSheetFormulaRegressionResultMatrix,
  webSheetFormulaRegressionStatsRow,
  webSheetFormulaRegularizedBeta,
  webSheetFormulaRegularizedGammaP,
  webSheetFormulaRoman,
  webSheetFormulaRoundAwayToParity,
  webSheetFormulaRoundToDecimals,
  webSheetFormulaRowKey,
  webSheetFormulaRri,
  webSheetFormulaScalar,
  webSheetFormulaSecurityBasis,
  webSheetFormulaSecurityDateSet,
  webSheetFormulaSecurityDuration,
  webSheetFormulaSecurityFrequency,
  webSheetFormulaSecurityPrice,
  webSheetFormulaSecurityYearFraction,
  webSheetFormulaSecurityYield,
  webSheetFormulaSequenceXRows,
  webSheetFormulaSeriesSum,
  webSheetFormulaSheetIndexFromFormulaSheetId,
  webSheetFormulaSignedBaseRange,
  webSheetFormulaSkew,
  webSheetFormulaSln,
  webSheetFormulaSplitTextByDelimiters,
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
  webSheetFormulaTbillDays,
  webSheetFormulaTbillEq,
  webSheetFormulaTbillPrice,
  webSheetFormulaTbillYield,
  webSheetFormulaText,
  webSheetFormulaTextDelimiters,
  webSheetFormulaTimeSerial,
  webSheetFormulaTimeUnit,
  webSheetFormulaTimeValue,
  webSheetFormulaTransposeMatrix,
  webSheetFormulaTrigReciprocal,
  webSheetFormulaTrimMean,
  webSheetFormulaUnwrapAstParenthesis,
  webSheetFormulaValidDateParts,
  webSheetFormulaVariance,
  webSheetFormulaVarianceFromArgs,
  webSheetFormulaVdb,
  webSheetFormulaWeekStartForReturnType,
  webSheetFormulaWeekdayValue,
  webSheetFormulaWeekendMask,
  webSheetFormulaWeeknumValue,
  webSheetFormulaWeibullDist,
  webSheetFormulaWildcardRegex,
  webSheetFormulaWorkbookFileName,
  webSheetFormulaWorkday,
  webSheetFormulaXirr,
  webSheetFormulaXlookupMatrix,
  webSheetFormulaXlookupResult,
  webSheetFormulaXnpv,
  webSheetFormulaYearFrac,
  webSheetFormulaYieldDisc,
  webSheetFormulaYieldMat,
  webSheetFormulaZTest,
};
