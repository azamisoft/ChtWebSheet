// Feature module; shared bindings are imported explicitly.
import {
  AUTO_FILTER_HEADER_SCAN_LIMIT,
  AUTO_FILTER_VALUE_LIMIT,
  SORT_CUSTOM_LISTS,
  sheetUsedRangeCacheBySheet,
  state,
} from "../../app-core.js";
import {
  formatQueryTimestamp,
  openExistingConnectionsDialog,
  queryConnectionTypeLabel,
  queryRefreshIntervalMinutes,
  refreshQueriesConnectionsPane,
  registerQueryInDataModel,
  registerQueryTableLoad,
  workbookQueries,
  workbookQueryKey,
  workbookQueryRowsForImport,
  writeWorkbookQueryRows,
} from "./connections-and-queries.js";
import {
  advancedFilterMatchingRows,
  autoFilterCellValue,
  autoFilterColorMenuHtml,
  autoFilterConditionMenuHtml,
  autoFilterDateSerial,
  autoFilterOperatorUsesBlankValue,
  autoFilterUniqueValues,
  autoFilterValueMatchesCriterion,
  closeAutoFilterMenu,
  escapeRegExp,
  isAutoFilterHeaderCell,
  normalizeAutoFilterCriterion,
  syncAutoFilterMenuState,
  updateAutoFilterSelectAllState,
} from "./filter-and-text-tools.js";
import {
  odataContentTypeFromUrl,
  odataImportSourcesFromSource,
  queryCredentialAuthOptions,
  requestHeadersForImportAuth,
  storeQueryCredential,
  webImportSourcesFromSource,
} from "./query-transforms.js";
import {
  decodeWorkbookRangeOrName,
} from "./validation-and-consolidation.js";
import {
  closeDataDialog,
} from "../drawing/charts-and-insertion.js";
import {
  mergesIntersectingRange,
  rangeContains,
} from "../editing/autofill.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  fallbackSingleCellRangeForSheet,
  getDisplayForCell,
  isCellInRange,
  isSingleCellRange,
  materializeSelectionRangeForSheet,
  rangeToLabel,
  restoreGroupedSheetsAfterFailure,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  cellHasFill,
  cellHasFontColor,
  rangesEqual,
  scrollRangeIntoView,
  selectionToolbarPoint,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  cellEffectiveCss,
  cloneCellModel,
  cloneJsonLikeValue,
  defaultColumnModel,
  defaultRowModel,
  pushHistory,
  writeClonedCellContent,
} from "../editing/structure-and-history.js";
import {
  forEachCellInRange,
  rangesIntersect,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  cellKey,
  columnName,
  escapeAttr,
  escapeHtml,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  applyExcelTableStyle,
  commitDataOperation,
  ensureSheetSize,
  parseCellKey,
  selectedDataRange,
} from "../ui/ribbon-and-links.js";
import {
  ensureSheetProtectionPermission,
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
  buildMergeMaps,
  cellDisplayText,
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  autoFilterStdDevCount,
  rangeRectEquals,
  tableAutoFilterRange,
} from "../workbook/tables-and-validation.js";


function applyQueryPropertiesDialog(dialog) {
  const { sheetIndex, queryId } = state.dataDialog || {};
  const returnToExistingConnectionsKey = state.dataDialog?.returnToExistingConnectionsKey || "";
  const sheet = state.model?.sheets?.[sheetIndex];
  const query = sheet?.queries?.find((item) => item.id === queryId);
  if (!sheet || !query) return;
  const name = String(dialog.querySelector("[data-data-field='name']")?.value || "").trim();
  if (!name) {
    showMessageBox("接続名を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const duplicate = workbookQueries().some(({ sheetIndex: itemSheetIndex, query: item }) =>
    item.id !== queryId &&
    itemSheetIndex !== undefined &&
    String(item.name || "").trim().toLocaleLowerCase("ja-JP") === name.toLocaleLowerCase("ja-JP"),
  );
  if (duplicate) {
    showMessageBox("同じ名前の接続が既に存在します。別の名前を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const refreshEveryEnabled = Boolean(dialog.querySelector("[data-data-field='refreshEveryEnabled']")?.checked);
  const rawInterval = Number(dialog.querySelector("[data-data-field='refreshIntervalMinutes']")?.value);
  if (refreshEveryEnabled && (!Number.isFinite(rawInterval) || rawInterval < 1 || rawInterval > 32767)) {
    showMessageBox("更新間隔には 1 から 32767 までの整数を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const authChanges = queryAuthenticationChangesFromPropertiesDialog(dialog, query);
  if (!authChanges.valid) {
    showMessageBox(authChanges.message, { title: "Cht WebSheet" });
    return;
  }
  const definitionChanges = queryDefinitionChangesFromPropertiesDialog(dialog, query, sheet);
  if (!definitionChanges.valid) {
    showMessageBox(definitionChanges.message, { title: "Cht WebSheet" });
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  query.name = name;
  query.description = dialog.querySelector("[data-data-field='description']")?.value || "";
  query.refreshWithAll = Boolean(dialog.querySelector("[data-data-field='refreshWithAll']")?.checked);
  query.backgroundRefresh = Boolean(dialog.querySelector("[data-data-field='backgroundRefresh']")?.checked);
  query.refreshOnOpen = Boolean(dialog.querySelector("[data-data-field='refreshOnOpen']")?.checked);
  query.refreshEveryEnabled = refreshEveryEnabled;
  query.refreshIntervalMinutes = Math.max(1, Math.min(32767, Math.round(rawInterval || queryRefreshIntervalMinutes(query))));
  query.removeDataBeforeSave = Boolean(dialog.querySelector("[data-data-field='removeDataBeforeSave']")?.checked);
  applyQueryDefinitionChanges(query, definitionChanges);
  applyQueryAuthenticationChanges(query, authChanges);
  const nextExistingConnectionKey = returnToExistingConnectionsKey ? workbookQueryKey(sheetIndex, query.id) : "";
  closeDataDialog();
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: query.range || null,
    selectionAfter: query.range || null,
    label: `${query.name || "Query"} プロパティ`,
  });
  if (sheetIndex === state.activeSheetIndex) renderSheet();
  refreshQueriesConnectionsPane();
  setStatus(`${query.name || "Query"} の接続プロパティを更新しました。`);
  scheduleQueryRefreshTimers();
  scheduleWorkbookPersistence();
  if (nextExistingConnectionKey) openExistingConnectionsDialog(nextExistingConnectionKey);
}



function queryDefinitionChangesFromPropertiesDialog(dialog, query, sheet) {
  const connectionString = String(dialog.querySelector("[data-data-field='connectionString']")?.value || "").trim();
  const commandType = String(dialog.querySelector("[data-data-field='commandType']")?.value || "default").trim() || "default";
  const commandText = String(dialog.querySelector("[data-data-field='commandText']")?.value || "").trim();
  const sourceItemName = String(dialog.querySelector("[data-data-field='sourceItemName']")?.value || "").trim();
  const parsed = parseQueryConnectionString(query, connectionString, commandText, sheet);
  if (["sqlServer", "access"].includes(query?.sourceType) && commandType === "sql" && !commandText) {
    return { valid: false, message: "SQL コマンドを入力してください。" };
  }
  if (["web", "odata", "json", "xml"].includes(query?.sourceType) && !parsed.sourceUrl && !commandText && !query?.sourceText) {
    return { valid: false, message: "接続文字列またはコマンド テキストに URL を入力してください。" };
  }
  return {
    valid: true,
    connectionFile: String(dialog.querySelector("[data-data-field='connectionFile']")?.value || ""),
    alwaysUseConnectionFile: Boolean(dialog.querySelector("[data-data-field='alwaysUseConnectionFile']")?.checked),
    failIfConnectionFileMissing: Boolean(dialog.querySelector("[data-data-field='failIfConnectionFileMissing']")?.checked),
    connectionString,
    commandType,
    commandText,
    sourceItemName,
    parsed,
  };
}



function applyQueryDefinitionChanges(query, changes) {
  if (!query || !changes?.valid) return;
  query.connectionFile = changes.connectionFile || "";
  query.alwaysUseConnectionFile = Boolean(changes.alwaysUseConnectionFile);
  query.failIfConnectionFileMissing = Boolean(changes.failIfConnectionFileMissing);
  query.connectionString = changes.connectionString || "";
  query.commandType = changes.commandType || "default";
  query.commandText = changes.commandText || "";
  if (changes.sourceItemName || query.sourceItemName) query.sourceItemName = changes.sourceItemName || "";
  const parsed = changes.parsed || {};
  if (parsed.sourceUrl !== undefined) query.sourceUrl = parsed.sourceUrl;
  if (parsed.sourceName !== undefined) query.sourceName = parsed.sourceName;
  if (parsed.sourceItemName !== undefined && !changes.sourceItemName) query.sourceItemName = parsed.sourceItemName;
  if (parsed.sourceSheetName !== undefined) query.sourceSheetName = parsed.sourceSheetName;
  if (parsed.sourceRange !== undefined) query.sourceRange = parsed.sourceRange;
  if (["web", "odata", "json", "xml"].includes(query.sourceType) && parsed.sourceUrl && !query.sourceName) {
    query.sourceName = parsed.sourceUrl;
  }
  query.lastRefreshStatus = "";
  query.refreshError = "";
}



function parseQueryConnectionString(query, connectionString, commandText = "", sheet = null) {
  const sourceType = query?.sourceType || "";
  const text = String(connectionString || "").trim();
  const command = String(commandText || "").trim();
  const prefixed = parsePrefixedConnectionString(text);
  const payload = prefixed.payload || text;
  if (["web", "odata", "json", "xml"].includes(sourceType)) {
    const url = connectionStringUrlForSourceType(sourceType, prefixed, payload) || firstUrlLikeText(command);
    const itemName = querySourceItemNameFromCommandText(command, url);
    return {
      ...(url ? { sourceUrl: url, sourceName: url } : {}),
      ...(itemName ? { sourceItemName: itemName } : {}),
    };
  }
  if (sourceType === "textCsv") {
    const sourceName = prefixed.prefix === "TEXT" ? textConnectionStringSource(payload) : payload || command || query?.sourceName || "";
    return sourceName ? { sourceName } : {};
  }
  if (sourceType === "workbook" || sourceType === "folder") {
    const sourceName = ["WORKBOOK", "FOLDER"].includes(prefixed.prefix) ? payload : payload || command || query?.sourceName || "";
    return sourceName ? { sourceName, sourceUrl: sourceName } : {};
  }
  if (sourceType === "sqlServer" || sourceType === "access") {
    const info = oleDbConnectionInfo(text);
    const sourceUrl = info.dataSource || payload || query?.sourceUrl || "";
    const sourceName = sourceType === "sqlServer" && info.database ? `${sourceUrl}/${info.database}` : sourceUrl;
    return {
      ...(sourceUrl ? { sourceUrl } : {}),
      ...(sourceName ? { sourceName } : {}),
      ...(command ? { sourceItemName: command } : {}),
    };
  }
  if (sourceType === "tableRange") {
    const rangeText = command || payload;
    const decoded = decodeWorkbookRangeOrName(rangeText, sheet || activeSheet());
    if (decoded) {
      const sourceSheet = state.model?.sheets?.[decoded.sheetIndex ?? state.activeSheetIndex];
      return {
        sourceSheetName: sourceSheet?.name || query?.sourceSheetName || "",
        sourceRange: { top: decoded.top, left: decoded.left, bottom: decoded.bottom, right: decoded.right },
      };
    }
  }
  return {};
}



function parsePrefixedConnectionString(connectionString) {
  const text = String(connectionString || "").trim();
  const match = text.match(/^([A-Za-z][\w ]*)\s*;\s*([\s\S]*)$/);
  if (!match) return { prefix: "", payload: text };
  return {
    prefix: match[1].replace(/\s+/g, "").toLocaleUpperCase("en-US"),
    payload: match[2] || "",
  };
}



function connectionStringUrlForSourceType(sourceType, prefixed, payload) {
  const prefix = prefixed?.prefix || "";
  const normalizedType = sourceType === "odata" ? "ODATA" : sourceType.toLocaleUpperCase("en-US");
  if (prefix && prefix !== normalizedType && !(sourceType === "web" && prefix === "WEB")) {
    return firstUrlLikeText(payload) || textConnectionStringSource(payload);
  }
  return firstUrlLikeText(payload) || textConnectionStringSource(payload);
}



function textConnectionStringSource(payload) {
  const text = String(payload || "").trim();
  if (!text) return "";
  return text.split(/;\s*(?:Origin|Delimiter|Format|Columns|Headers|Encoding|Locale|Data Source)\s*=/i)[0].trim();
}



function firstUrlLikeText(text) {
  const source = String(text || "").trim();
  const match = source.match(/(?:https?:|data:)[^\s"')]+/i);
  if (match) return match[0];
  if (/^(?:https?:|data:)/i.test(source)) return source;
  return "";
}



function querySourceItemNameFromCommandText(commandText, sourceUrl = "") {
  const command = String(commandText || "").trim();
  if (!command) return "";
  const source = String(sourceUrl || "").trim();
  if (source && command.startsWith(source)) {
    return command.slice(source.length).replace(/^\s*[-:]\s*/, "").trim();
  }
  return /^(?:https?:|data:)/i.test(command) ? "" : command;
}



function oleDbConnectionInfo(connectionString) {
  const dataSource =
    connectionStringValue(connectionString, "Data Source") ||
    connectionStringValue(connectionString, "Server") ||
    connectionStringValue(connectionString, "Address");
  const database =
    connectionStringValue(connectionString, "Initial Catalog") ||
    connectionStringValue(connectionString, "Database");
  return { dataSource, database };
}



function connectionStringValue(connectionString, key) {
  const pattern = new RegExp(`(?:^|;)\\s*${escapeRegExp(key)}\\s*=\\s*([^;]*)`, "i");
  return String(connectionString || "").match(pattern)?.[1]?.trim() || "";
}



function setConnectionStringValue(connectionString, key, value) {
  const text = String(connectionString || "").trim();
  const normalizedKey = String(key || "").trim();
  if (!normalizedKey) return text;
  const pattern = new RegExp(`(^|;)\\s*${escapeRegExp(normalizedKey)}\\s*=\\s*[^;]*`, "i");
  if (pattern.test(text)) {
    return text.replace(pattern, (match, separator = "") => `${separator}${normalizedKey}=${String(value || "")}`);
  }
  const separator = text && !text.endsWith(";") ? ";" : "";
  return `${text}${separator}${normalizedKey}=${String(value || "")}`;
}



function queryAuthenticationChangesFromPropertiesDialog(dialog, query) {
  if (!dialog?.querySelector("[name='queryAuthMode']")) return { valid: true, mode: query.sourceAuthMode || "anonymous" };
  const mode = dialog.querySelector("[name='queryAuthMode']:checked")?.value === "basic" ? "basic" : "anonymous";
  if (mode !== "basic") return { valid: true, mode: "anonymous" };
  const username = String(dialog.querySelector("[data-data-field='queryAuthUsername']")?.value || "").trim();
  const password = String(dialog.querySelector("[data-data-field='queryAuthPassword']")?.value || "");
  const existingCredential = query?.sourceCredentialKey ? state.queryCredentials?.get?.(query.sourceCredentialKey) : null;
  if (!username) {
    return { valid: false, message: "基本認証のユーザー名を入力してください。" };
  }
  if (!password && !existingCredential) {
    return { valid: false, message: "基本認証のパスワードを入力してください。" };
  }
  return {
    valid: true,
    mode: "basic",
    username,
    password,
    preserveExistingPassword: !password && Boolean(existingCredential),
  };
}



function applyQueryAuthenticationChanges(query, changes) {
  if (!query || !changes?.valid) return;
  const previousCredentialKey = query.sourceCredentialKey || "";
  if (changes.mode !== "basic") {
    if (previousCredentialKey) state.queryCredentials?.delete?.(previousCredentialKey);
    query.sourceAuthMode = "";
    query.sourceAuthUsername = "";
    query.sourceCredentialKey = "";
    return;
  }
  query.sourceAuthMode = "basic";
  query.sourceAuthUsername = changes.username || "";
  if (changes.preserveExistingPassword && previousCredentialKey) {
    const credential = state.queryCredentials?.get?.(previousCredentialKey);
    if (credential) credential.username = changes.username || credential.username || "";
    query.sourceCredentialKey = previousCredentialKey;
    return;
  }
  if (previousCredentialKey) state.queryCredentials?.delete?.(previousCredentialKey);
  query.sourceCredentialKey = storeQueryCredential({
    mode: "basic",
    username: changes.username,
    password: changes.password,
  });
}



function clearQueryPropertiesCredential(dialog) {
  const { sheetIndex, queryId } = state.dataDialog || {};
  const sheet = state.model?.sheets?.[sheetIndex];
  const query = sheet?.queries?.find((item) => item.id === queryId);
  if (!query?.sourceCredentialKey) return;
  state.queryCredentials?.delete?.(query.sourceCredentialKey);
  query.sourceCredentialKey = "";
  const status = dialog?.querySelector("[data-query-auth-session-status]");
  if (status) status.textContent = "資格情報をクリアしました。OK を押す前にパスワードを入力してください。";
  dialog?.querySelector("[data-data-field='queryAuthPassword']")?.focus({ preventScroll: true });
  const button = dialog?.querySelector("[data-data-dialog-action='query-properties-clear-credential']");
  if (button) {
    button.dataset.queryAuthCanClear = "false";
    button.disabled = true;
  }
  setStatus(`${query.name || "Query"} の資格情報をクリアしました。`);
}



function handleQueryPropertiesDialogSync(event) {
  const dialog = event.target.closest("[data-data-dialog='query-properties']");
  if (!dialog) return;
  const refreshEvery = Boolean(dialog.querySelector("[data-data-field='refreshEveryEnabled']")?.checked);
  const interval = dialog.querySelector("[data-data-field='refreshIntervalMinutes']");
  if (interval) interval.disabled = !refreshEvery;
  syncQueryAuthenticationProperties(dialog);
}



function syncQueryAuthenticationProperties(dialog) {
  const fields = dialog?.querySelector("[data-query-auth-basic-fields]");
  if (!fields) return;
  const isBasic = dialog.querySelector("[name='queryAuthMode']:checked")?.value === "basic";
  fields.classList.toggle("hidden", !isBasic);
  fields.querySelectorAll("input, button").forEach((control) => {
    if (control.matches("[data-data-dialog-action='query-properties-clear-credential']")) {
      control.disabled = !isBasic || control.dataset.queryAuthCanClear !== "true";
      return;
    }
    control.disabled = !isBasic;
  });
}



function detectDelimitedTextDelimiter(text, sourceName = "") {
  const name = String(sourceName).toLowerCase();
  if (name.endsWith(".tsv")) return "\t";
  const sample = String(text || "").split(/\r\n|\r|\n/).slice(0, 8).join("\n");
  const candidates = ["\t", ",", ";"];
  return candidates
    .map((delimiter) => ({ delimiter, count: (sample.match(new RegExp(delimiter === "\t" ? "\\t" : `\\${delimiter}`, "g")) || []).length }))
    .sort((a, b) => b.count - a.count)[0]?.delimiter || ",";
}



function parseDelimitedRows(text, delimiter = ",") {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  const source = String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === '"') {
      if (inQuotes && source[index + 1] === '"') {
        field += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (!inQuotes && char === delimiter) {
      row.push(field);
      field = "";
      continue;
    }
    if (!inQuotes && char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      continue;
    }
    field += char;
  }
  row.push(field);
  rows.push(row);
  return rows;
}



async function refreshWorkbookData() {
  if (!state.model) return;
  const targets = [];
  (state.model.sheets || []).forEach((sheet, sheetIndex) => {
    (sheet.queries || []).forEach((query) => {
      if (query.refreshWithAll !== false) {
        targets.push({ sheet, sheetIndex, query });
      }
    });
  });
  if (!targets.length) {
    setStatus("更新する接続はありません。");
    return;
  }
  const targetSheetIndexes = targets.map((target) => target.sheetIndex);
  const sheetsBefore = captureSheetsForHistory(targetSheetIndexes);
  const failures = [];
  setStatus(`${targets.length} 件の接続を更新しています...`);
  for (const { sheet, query } of targets) {
    try {
      await refreshWorkbookQueryWithRemoteSource(sheet, query);
      if (query.addToDataModel) {
        const sheetIndex = state.model?.sheets?.indexOf(sheet);
        if (sheetIndex >= 0) registerQueryInDataModel(sheetIndex, query);
      }
    } catch (error) {
      console.error(error);
      failures.push({ query, error });
    }
  }
  state.hf = buildFormulaEngine(state.model);
  reapplyAutoFilterForSheetRefs(targets, { includeSort: true });
  state.hf = buildFormulaEngine(state.model);
  const sheetsAfter = captureSheetsForHistory(targetSheetIndexes);
  if (sheetsBefore.length) {
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter,
      selectionBefore: state.selectionRange || null,
      selectionAfter: state.selectionRange || null,
      label: "すべて更新",
    });
  }
  renderSheet();
  updateFormulaBarForSelection();
  refreshQueriesConnectionsPane();
  const successCount = targets.length - failures.length;
  if (failures.length) {
    setStatus(`${successCount} 件の接続を更新しました。${failures.length} 件は更新できませんでした。`);
    showMessageBox(
      failures
        .map(({ query, error }) => `${query?.name || "Query"}: ${readableRefreshError(error)}`)
        .join("\n"),
      { title: "Cht WebSheet" },
    );
  } else {
    setStatus(`${targets.length} 件の接続を更新しました。`);
  }
  scheduleWorkbookPersistence();
  scheduleQueryRefreshTimers();
}



function refreshWorkbookQueriesOnOpen() {
  if (!state.model) return 0;
  const targets = [];
  (state.model.sheets || []).forEach((sheet, sheetIndex) => {
    (sheet.queries || []).forEach((query) => {
      if (query.refreshOnOpen) targets.push({ sheet, sheetIndex, query });
    });
  });
  if (!targets.length) return 0;
  const loadVersion = state.workbookLoadVersion;
  Promise.resolve().then(() => refreshWorkbookQueriesOnOpenAsync(targets, loadVersion));
  return targets.length;
}



async function refreshWorkbookQueriesOnOpenAsync(targets, loadVersion) {
  const currentTargets = (targets || []).filter((target) => queryRefreshTargetStillCurrent(target, loadVersion));
  if (!currentTargets.length) return;
  const failures = [];
  let successCount = 0;
  setStatus(`${currentTargets.length} 件の接続を開くときに更新しています...`);
  refreshQueriesConnectionsPane();
  for (const target of currentTargets) {
    if (!queryRefreshTargetStillCurrent(target, loadVersion)) return;
    const { sheet, sheetIndex, query } = target;
    try {
      await refreshWorkbookQueryWithRemoteSource(sheet, query);
      if (query.addToDataModel) registerQueryInDataModel(sheetIndex, query);
      successCount += 1;
    } catch (error) {
      console.error(error);
      failures.push({ query, error });
    }
  }
  if (!queryRefreshTargetStillCurrent(currentTargets[0], loadVersion)) return;
  state.hf = buildFormulaEngine(state.model);
  reapplyAutoFilterForSheetRefs(currentTargets, { includeSort: true });
  state.hf = buildFormulaEngine(state.model);
  renderSheet();
  updateFormulaBarForSelection();
  refreshQueriesConnectionsPane();
  scheduleWorkbookPersistence();
  scheduleQueryRefreshTimers();
  if (failures.length) {
    setStatus(`${successCount} 件の接続を開くときに更新しました。${failures.length} 件は更新できませんでした。`);
    showMessageBox(
      failures
        .map(({ query, error }) => `${query?.name || "Query"}: ${readableRefreshError(error)}`)
        .join("\n"),
      { title: "Cht WebSheet" },
    );
    return;
  }
  setStatus(`${successCount} 件の接続を開くときに更新しました。`);
}



function queryRefreshTargetStillCurrent(target, loadVersion) {
  if (!target || !state.model) return false;
  if (Number.isFinite(Number(loadVersion)) && Number(loadVersion) !== Number(state.workbookLoadVersion)) return false;
  return state.model.sheets?.[target.sheetIndex] === target.sheet && target.sheet?.queries?.includes(target.query);
}



function reapplyAutoFilterForSheetRefs(refs, options = {}) {
  const unique = new Map();
  (refs || []).forEach((ref) => {
    if (!ref?.sheet) return;
    const sheetIndex = Number.isFinite(Number(ref.sheetIndex))
      ? Number(ref.sheetIndex)
      : state.model?.sheets?.indexOf(ref.sheet);
    const key = sheetIndex >= 0 ? `i:${sheetIndex}` : `n:${ref.sheet.name || unique.size}`;
    if (!unique.has(key)) unique.set(key, { sheet: ref.sheet, sheetIndex });
  });
  unique.forEach((ref) => {
    reapplyAutoFilter({
      silent: true,
      includeSort: Boolean(options.includeSort),
      sheet: ref.sheet,
      sheetIndex: ref.sheetIndex,
    });
  });
}



function clearQueryRefreshTimers() {
  (state.queryRefreshTimers || []).forEach((entry) => {
    if (entry?.timer) window.clearTimeout(entry.timer);
  });
  state.queryRefreshTimers = [];
}



function scheduleQueryRefreshTimers() {
  clearQueryRefreshTimers();
  if (!state.model) return;
  (state.model.sheets || []).forEach((sheet, sheetIndex) => {
    (sheet.queries || []).forEach((query) => {
      if (!query?.refreshEveryEnabled) return;
      const delay = queryRefreshIntervalMinutes(query) * 60 * 1000;
      const queryId = query.id;
      const timer = window.setTimeout(() => refreshScheduledQuery(sheetIndex, queryId), delay);
      state.queryRefreshTimers.push({ sheetIndex, queryId, timer });
    });
  });
}



async function refreshScheduledQuery(sheetIndex, queryId) {
  const sheet = state.model?.sheets?.[sheetIndex];
  const query = sheet?.queries?.find((item) => item.id === queryId);
  if (!sheet || !query || !query.refreshEveryEnabled) {
    scheduleQueryRefreshTimers();
    return;
  }
  try {
    await refreshWorkbookQueryWithRemoteSource(sheet, query);
  } catch (error) {
    console.error(error);
    setStatus(`${query.name || "Query"} を自動更新できませんでした。`);
  }
  state.hf = buildFormulaEngine(state.model);
  reapplyAutoFilter({ silent: true, includeSort: true, sheet, sheetIndex });
  state.hf = buildFormulaEngine(state.model);
  if (sheetIndex === state.activeSheetIndex) renderSheet();
  refreshQueriesConnectionsPane();
  if (query.lastRefreshStatus !== "failed") setStatus(`${query.name || "Query"} を自動更新しました。`);
  scheduleWorkbookPersistence();
  scheduleQueryRefreshTimers();
}



async function refreshWorkbookQueryWithRemoteSource(sheet, query) {
  markWorkbookQueryRefreshing(query);
  refreshQueriesConnectionsPane();
  try {
    const preflightError = queryRefreshPreflightError(query);
    if (preflightError) throw new Error(preflightError);
    await refreshWorkbookQueryRemoteSource(query);
    refreshWorkbookQuery(sheet, query);
  } catch (error) {
    if (!["failed", "cancelled"].includes(query?.lastRefreshStatus)) markWorkbookQueryFailed(query, error);
    throw error;
  } finally {
    refreshQueriesConnectionsPane();
  }
}



function queryRefreshPreflightError(query) {
  if (!query) return "接続が見つかりません。";
  if (query.alwaysUseConnectionFile && query.failIfConnectionFileMissing && !String(query.connectionFile || "").trim()) {
    return "接続ファイルが指定されていません。";
  }
  if (["web", "odata", "json", "xml"].includes(query.sourceType) && !query.sourceUrl && !query.sourceText) {
    return "データ ソースの URL が指定されていません。";
  }
  return "";
}



function markWorkbookQueryRefreshing(query) {
  if (!query) return;
  query.lastRefreshStarted = new Date().toISOString();
  query.lastRefreshStatus = "refreshing";
  query.refreshError = "";
}



function markWorkbookQuerySucceeded(query) {
  if (!query) return;
  const endedAt = Date.now();
  query.lastRefreshed = new Date(endedAt).toISOString();
  query.lastRefreshDurationMs = queryRefreshDurationMs(query, endedAt);
  query.lastRefreshStarted = "";
  query.refreshCount = Math.max(0, Number(query.refreshCount) || 0) + 1;
  query.lastRefreshStatus = "success";
  query.refreshError = "";
}



function markWorkbookQueryFailed(query, error) {
  if (!query) return;
  query.lastRefreshDurationMs = queryRefreshDurationMs(query);
  query.lastRefreshStarted = "";
  query.lastRefreshStatus = error?.name === "AbortError" ? "cancelled" : "failed";
  query.refreshError = readableRefreshError(error);
}



async function refreshWorkbookQueryRemoteSource(query) {
  if (!query || !["web", "odata", "json", "xml"].includes(query.sourceType) || !query.sourceUrl) return;
  if (!/^(?:https?:|data:)/i.test(query.sourceUrl)) return;
  const controller = typeof AbortController === "function" ? new AbortController() : null;
  const job = {
    queryId: query.id || "",
    name: query.name || "Query",
    startedAt: new Date().toISOString(),
    controller,
  };
  state.activeQueryRefreshes = Array.isArray(state.activeQueryRefreshes) ? state.activeQueryRefreshes : [];
  state.activeQueryRefreshes.push(job);
  query.lastRefreshStatus = "refreshing";
  query.refreshError = "";
  try {
    const headers = query.sourceType === "odata"
      ? { accept: "application/json, application/xml, text/xml, text/csv, */*" }
      : query.sourceType === "json"
        ? { accept: "application/json, text/json, */*" }
        : query.sourceType === "xml"
          ? { accept: "application/xml, text/xml, */*" }
          : {};
    const auth = queryCredentialAuthOptions(query);
    const requestHeaders = requestHeadersForImportAuth(auth, headers);
    const response = await fetch(query.sourceUrl, { headers: requestHeaders, ...(controller ? { signal: controller.signal } : {}) });
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`.trim());
    const contentType = response.headers.get("content-type") || (query.sourceType === "odata" ? odataContentTypeFromUrl(query.sourceUrl) : "");
    const sourceText = await response.text();
    const sources = query.sourceType === "odata"
      ? odataImportSourcesFromSource(sourceText, contentType, query.sourceUrl)
      : webImportSourcesFromSource(sourceText, contentType, query.sourceUrl);
    const selected = querySourceForRefresh(query, sources);
    if (!selected?.rows?.length) throw new Error("読み込める表データが見つかりません。");
    query.sourceText = sourceText;
    query.contentType = contentType;
    query.sourceItemIndex = selected.itemIndex ?? 0;
    query.sourceItemName = selected.name || query.sourceItemName || "";
    query.sourceItemKind = selected.kind || query.sourceItemKind || "document";
	    query.refreshError = "";
	  } catch (error) {
	    markWorkbookQueryFailed(query, error);
	    throw error;
  } finally {
    state.activeQueryRefreshes = (state.activeQueryRefreshes || []).filter((entry) => entry !== job);
  }
}



function querySourceForRefresh(query, sources = []) {
  const normalizedIndex = Math.max(0, Math.round(Number(query?.sourceItemIndex) || 0));
  return (
    sources.find((item) => Number(item.itemIndex) === normalizedIndex) ||
    sources.find((item) => item.name && item.name === query?.sourceItemName) ||
    sources[normalizedIndex] ||
    sources[0] ||
    null
  );
}



function readableRefreshError(error) {
  if (!error) return "不明なエラー";
  if (error.name === "AbortError") return "更新はキャンセルされました。";
  return String(error.message || error || "不明なエラー");
}



function cancelWorkbookRefreshes() {
  const jobs = [...(state.activeQueryRefreshes || [])];
  jobs.forEach((job) => job?.controller?.abort?.());
  setStatus(jobs.length ? `${jobs.length} 件の更新をキャンセルしました。` : "実行中の更新はありません。");
  refreshQueriesConnectionsPane();
}



function openRefreshStatusDialog() {
  if (!state.model) return;
  state.dataDialog = { type: "refresh-status" };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = refreshStatusDialogHtml();
  document.body.appendChild(overlay);
  overlay.querySelector("[data-data-dialog-action='refresh-status-close']")?.focus({ preventScroll: true });
}



function refreshStatusDialogHtml() {
  const rows = workbookQueries();
  const body = rows.length
    ? rows.map(({ sheet, sheetIndex, query }) => refreshStatusRowHtml(sheet, sheetIndex, query)).join("")
    : `<tr><td colspan="8">このブックには接続がありません。</td></tr>`;
  const activeCount = activeWorkbookRefreshCount();
  return `
    <div class="excel-mini-dialog data-dialog refresh-status-dialog" role="dialog" aria-modal="true" aria-labelledby="refreshStatusTitle" data-data-dialog="refresh-status">
      <div class="excel-mini-titlebar" id="refreshStatusTitle">更新の状態</div>
      <div class="excel-mini-body data-dialog-body">
	        <div class="data-dialog-note">${activeCount ? `${activeCount} 件の更新が実行中です。` : "現在実行中の更新はありません。"}</div>
	        <div class="refresh-status-table-wrap">
	          <table class="refresh-status-table">
	            <thead>
	              <tr><th>名前</th><th>種類</th><th>場所</th><th>状態</th><th>最終更新</th><th>経過</th><th>更新回数</th><th>操作</th></tr>
	            </thead>
	            <tbody>${body}</tbody>
	          </table>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="refresh-status-refresh-all">すべて更新</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="refresh-status-cancel" ${activeCount ? "" : "disabled"}>更新のキャンセル</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="refresh-status-close">閉じる</button>
      </div>
    </div>`;
}



function activeWorkbookRefreshCount() {
  const activeKeys = new Set();
  (state.activeQueryRefreshes || []).forEach((job) => {
    if (job.queryId) activeKeys.add(`query:${job.queryId}`);
  });
  workbookQueries().forEach(({ sheetIndex, query }) => {
    if (query.lastRefreshStatus === "refreshing") activeKeys.add(query.id ? `query:${query.id}` : workbookQueryKey(sheetIndex, query.id));
  });
  return activeKeys.size;
}



function refreshStatusRowHtml(sheet, sheetIndex, query) {
  const active = query.lastRefreshStatus === "refreshing" || (state.activeQueryRefreshes || []).some((job) => job.queryId && job.queryId === query.id);
  const status = active
    ? "更新中"
    : query.lastRefreshStatus === "failed"
      ? `失敗: ${query.refreshError || "更新できませんでした"}`
      : query.lastRefreshStatus === "cancelled"
        ? "キャンセル済み"
        : query.lastRefreshed
          ? "完了"
          : "未更新";
  const range = query.range ? `${sheet.name}!${rangeToLabel(query.range)}` : sheet.name;
  return `
	    <tr data-refresh-query-key="${escapeAttr(workbookQueryKey(sheetIndex, query.id))}">
	      <td>${escapeHtml(query.name || "Query")}</td>
	      <td>${escapeHtml(queryConnectionTypeLabel(query))}</td>
	      <td>${escapeHtml(range)}</td>
	      <td>${escapeHtml(status)}</td>
	      <td>${escapeHtml(formatQueryTimestamp(query.lastRefreshed || ""))}</td>
	      <td>${escapeHtml(queryRefreshDurationLabel(query, active))}</td>
	      <td>${escapeHtml(String(Math.max(0, Number(query.refreshCount) || 0)))}</td>
	      <td class="refresh-status-actions">
	        <button class="excel-mini-button" type="button" data-data-dialog-action="refresh-status-select" data-query-sheet-index="${sheetIndex}" data-query-id="${escapeAttr(query.id)}" ${query.range ? "" : "disabled"}>移動</button>
	        <button class="excel-mini-button" type="button" data-data-dialog-action="refresh-status-properties" data-query-sheet-index="${sheetIndex}" data-query-id="${escapeAttr(query.id)}">プロパティ</button>
	      </td>
	    </tr>`;
}



function queryRefreshDurationLabel(query, active = false) {
  const startedAt = new Date(query?.lastRefreshStarted || "").getTime();
  if (active && Number.isFinite(startedAt)) {
    return `${Math.max(0, Math.round((Date.now() - startedAt) / 1000))} 秒`;
  }
  const duration = Number(query?.lastRefreshDurationMs);
  if (!Number.isFinite(duration) || duration <= 0) return "";
  if (duration < 1000) return `${Math.round(duration)} ms`;
  return `${(duration / 1000).toFixed(duration < 10000 ? 1 : 0)} 秒`;
}



function queryRefreshDurationMs(query, endedAt = Date.now()) {
  const startedAt = new Date(query?.lastRefreshStarted || "").getTime();
  return Number.isFinite(startedAt) ? Math.max(0, endedAt - startedAt) : 0;
}



function refreshWorkbookQuery(sheet, query) {
  if (!sheet || !query) return;
  if (query.sourceType === "tableRange" && query.range) {
    if (query.transformSteps?.length) {
      const rows = workbookQueryRowsForImport({ sheet, sheetIndex: state.model?.sheets?.indexOf(sheet) ?? state.activeSheetIndex, query });
      if (rows.height && rows.width) {
        const targetRange = {
          ...query.range,
          bottom: query.range.top + rows.height - 1,
          right: query.range.left + rows.width - 1,
        };
        ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
        writeWorkbookQueryRows(sheet, rows, targetRange.top, targetRange.left);
        query.range = targetRange;
        registerQueryTableLoad(sheet, targetRange, query);
      }
    } else {
      ensureSheetSize(sheet, query.range.bottom, query.range.right);
      applyExcelTableStyle(sheet, query.range);
      sheet.tables = Array.isArray(sheet.tables) ? sheet.tables : [];
      const table = sheet.tables.find((item) => item.range && rangesEqual(item.range, query.range));
      if (table) {
        table.showHeaderRow = query.hasHeader !== false;
        table.showFilterButton = true;
      }
    }
  }
	  if (query.sourceType === "textCsv" && query.range && query.sourceText != null) {
	    const rows = workbookQueryRowsForImport({ sheet, sheetIndex: state.model?.sheets?.indexOf(sheet) ?? state.activeSheetIndex, query });
	    if (!rows.height || !rows.width) {
	      markWorkbookQuerySucceeded(query);
	      return;
	    }
    const targetRange = {
      ...query.range,
      bottom: query.range.top + rows.height - 1,
      right: query.range.left + rows.width - 1,
    };
    ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
    writeWorkbookQueryRows(sheet, rows, targetRange.top, targetRange.left);
    query.range = targetRange;
    registerQueryTableLoad(sheet, targetRange, query);
  }
	  if (["web", "json", "xml"].includes(query.sourceType) && query.range && query.sourceText != null) {
	    const rows = workbookQueryRowsForImport({ sheet, sheetIndex: state.model?.sheets?.indexOf(sheet) ?? state.activeSheetIndex, query });
	    if (!rows.height || !rows.width) {
	      markWorkbookQuerySucceeded(query);
	      return;
	    }
    const targetRange = {
      ...query.range,
      bottom: query.range.top + rows.height - 1,
      right: query.range.left + rows.width - 1,
    };
    ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
    writeWorkbookQueryRows(sheet, rows, targetRange.top, targetRange.left);
    query.range = targetRange;
    registerQueryTableLoad(sheet, targetRange, query);
  }
	  if (["odata", "workbook", "folder", "queryReference", "queryAppend", "queryMerge"].includes(query.sourceType) && query.range) {
	    const rows = workbookQueryRowsForImport({ sheet, sheetIndex: state.model?.sheets?.indexOf(sheet) ?? state.activeSheetIndex, query });
	    if (!rows.height || !rows.width) {
	      markWorkbookQuerySucceeded(query);
	      return;
	    }
    const targetRange = {
      ...query.range,
      bottom: query.range.top + rows.height - 1,
      right: query.range.left + rows.width - 1,
    };
    ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
    writeWorkbookQueryRows(sheet, rows, targetRange.top, targetRange.left);
    query.range = targetRange;
    registerQueryTableLoad(sheet, targetRange, query);
  }
	  markWorkbookQuerySucceeded(query);
	}



function sortSelectionByActiveColumn(direction) {
  const context = selectedDataRange("並べ替え");
  if (!context) return;
  if (!ensureSheetProtectionPermission(context.sheet, "allowSort", "保護されたシートでは並べ替えを使用できません。")) return;
  const active = state.selected?.sheetIndex === state.activeSheetIndex ? state.selected : null;
  const expandedRange = quickSortRangeFromSelection(context.sheet, context.range);
  if (sortSelectionByActiveColumnForSheetGroup(context, expandedRange, active, direction)) return;
  sortRangeByActiveColumn(context.sheet, expandedRange, active, direction);
}



function sortRangeByActiveColumn(sheet, range, active, direction) {
  if (!sheet || !range) return;
  const keyCol = active && active.col >= range.left && active.col <= range.right ? active.col : range.left;
  sortRangeRows(range, keyCol, direction, inferHeaderRow(sheet, range));
}



function sortSelectionByActiveColumnForSheetGroup(context, activeExpandedRange, activePoint, direction) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !context?.range || !activeExpandedRange) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowSort", "保護されたシートでは並べ替えを使用できません。")) return true;
    const selectionRange = materializeSelectionRangeForSheet(sheet, context.range, sheetIndex);
    if (!selectionRange) continue;
    const activeForSheet = activePoint
      ? { sheetIndex, row: activePoint.row, col: activePoint.col }
      : { sheetIndex, row: selectionRange.top, col: selectionRange.left };
    const range = quickSortRangeFromSelection(sheet, selectionRange, { activePoint: activeForSheet, sheetIndex });
    const keyCol = activeForSheet.col >= range.left && activeForSheet.col <= range.right ? activeForSheet.col : range.left;
    const hasHeader = inferHeaderRow(sheet, range);
    items.push({ sheetIndex, sheet, range, keyCol, hasHeader });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    const failed = items.find((item) => !sortRangeRowsByKeys(
      item.range,
      [{ col: item.keyCol, direction }],
      item.hasHeader,
      { sheet: item.sheet, commitHistory: false, quiet: true },
    ));
    if (failed) {
      restoreGroupedSheetsAfterFailure(sheetsBefore);
      showMessageBox(`${failed.sheet.name} の範囲は並べ替えできません。`, { title: "Cht WebSheet" });
      return true;
    }
    state.hf = buildFormulaEngine(state.model);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: context.range,
      selectionAfter: context.range,
      selectionBeforeRanges: activeSelectionRanges().map((range) => ({ ...range })),
      selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
      label: `${sheetIndexes.length} シート ${direction === "desc" ? "降順" : "昇順"}並べ替え`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートを${direction === "desc" ? "降順" : "昇順"}で並べ替えました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`並べ替えに失敗しました：${error.message || error}`);
  }
  return true;
}



function shouldShowSortExpansionWarning(sheet, selectedRange, expandedRange) {
  if (!sheet || !selectedRange || !expandedRange) return false;
  if (isSingleCellRange(selectedRange) || rangesEqual(selectedRange, expandedRange)) return false;
  if (!rangeContains(expandedRange, selectedRange)) return false;
  const selectedCells = (selectedRange.bottom - selectedRange.top + 1) * (selectedRange.right - selectedRange.left + 1);
  const expandedCells = (expandedRange.bottom - expandedRange.top + 1) * (expandedRange.right - expandedRange.left + 1);
  return expandedCells > selectedCells;
}



function quickSortRangeFromSelection(sheet, range, options = {}) {
  if (!sheet || !range) return range;
  const point = options.activePoint || (state.selected?.sheetIndex === state.activeSheetIndex
    ? state.selected
    : { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left });
  const region = currentRegionFromPoint(sheet, point, options.sheetIndex ?? point.sheetIndex ?? state.activeSheetIndex);
  if (!region) return range;
  if (isSingleCellRange(range)) return region;
  if (rangeContains(region, range) && !rangesEqual(region, range)) return region;
  return range;
}



function currentRegionFromPoint(sheet, point, sheetIndex = state.activeSheetIndex) {
  const used = sheetUsedRange(sheet);
  if (!used || !point) return null;
  if (point.row < used.top || point.row > used.bottom || point.col < used.left || point.col > used.right) {
    return { sheetIndex, top: point.row, bottom: point.row, left: point.col, right: point.col };
  }
  let region = { sheetIndex, top: point.row, bottom: point.row, left: point.col, right: point.col };
  let changed = true;
  while (changed) {
    changed = false;
    if (region.left > used.left && rangeColumnHasData(sheet, region.left - 1, region.top, region.bottom)) {
      region.left -= 1;
      changed = true;
    }
    if (region.right < used.right && rangeColumnHasData(sheet, region.right + 1, region.top, region.bottom)) {
      region.right += 1;
      changed = true;
    }
    if (region.top > used.top && rangeRowHasData(sheet, region.top - 1, region.left, region.right)) {
      region.top -= 1;
      changed = true;
    }
    if (region.bottom < used.bottom && rangeRowHasData(sheet, region.bottom + 1, region.left, region.right)) {
      region.bottom += 1;
      changed = true;
    }
  }
  return region;
}



function rangeRowHasData(sheet, row, left, right) {
  for (let col = left; col <= right; col += 1) {
    if (cellHasDataForRegion(sheet, row, col)) return true;
  }
  return false;
}



function rangeColumnHasData(sheet, col, top, bottom) {
  for (let row = top; row <= bottom; row += 1) {
    if (cellHasDataForRegion(sheet, row, col)) return true;
  }
  return false;
}



function cellHasDataForRegion(sheet, row, col) {
  const cell = sheet?.cells?.[cellKey(row, col)];
  return Boolean(cell?.formula || String(cellDisplayText(cell)).trim());
}



function rangeHasUserData(sheet, range, ignoreRange = null) {
  let found = false;
  forEachCellInRange(range, (row, col) => {
    if (found) return;
    if (ignoreRange && isCellInRange(ignoreRange, row, col)) return;
    if (cellHasDataForRegion(sheet, row, col)) found = true;
  });
  return found;
}



function showDataOverwriteConfirmation(onConfirm) {
  showMessageBox("貼り付け先のセルには既にデータがあります。置き換えますか?", {
    title: "Cht WebSheet",
    tone: "warning",
    buttons: [
      { label: "OK", result: "ok", callback: onConfirm },
      { label: "キャンセル", result: "cancel" },
    ],
    cancelResult: "cancel",
  });
}



function openSortDialog(initialRange = null, skipExpansionWarning = false) {
  const context = initialRange
    ? { sheet: activeSheet(), range: initialRange }
    : selectedDataRange("並べ替え");
  if (!context) return;
  if (!ensureSheetProtectionPermission(context.sheet, "allowSort", "保護されたシートでは並べ替えを使用できません。")) return;
  const { sheet } = context;
  const range = skipExpansionWarning || !isSingleCellRange(context.range)
    ? context.range
    : quickSortRangeFromSelection(sheet, context.range);
  if (!skipExpansionWarning && isSingleCellRange(context.range) && shouldShowSortExpansionWarning(sheet, context.range, range)) {
    showMessageBox("選択範囲の周囲にデータが見つかりました。並べ替えの対象をどうしますか?", {
      title: "並べ替えの前に",
      tone: "warning",
      buttons: [
        { label: "選択範囲を拡張する(E)", result: "expand", callback: () => openSortDialog(range, true) },
        { label: "現在選択されている範囲を並べ替える(C)", result: "current", callback: () => openSortDialog(context.range, true) },
        { label: "キャンセル", result: "cancel" },
      ],
      cancelResult: "cancel",
    });
    return;
  }
  const storedSortState = sortStateForDialog(sheet, range);
  const dialogRange = storedSortState?.range ? { ...storedSortState.range, sheetIndex: storedSortState.range.sheetIndex ?? state.activeSheetIndex } : range;
  const orientation = storedSortState?.orientation === "columns" ? "columns" : "rows";
  const hasHeader = storedSortState ? Boolean(storedSortState.hasHeader) : inferHeaderRow(sheet, dialogRange);
  const caseSensitive = Boolean(storedSortState?.caseSensitive);
  const sortLevels = storedSortState?.keys?.length
    ? sortDialogLevelsFromSortState(storedSortState, orientation)
    : [defaultSortDialogLevel(dialogRange, orientation)];
  state.dataDialog = {
    type: "sort",
    range: dialogRange,
    sortOrientation: orientation,
    sortLevels,
    sortActiveLevel: 0,
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog sort-data-dialog" role="dialog" aria-modal="true" aria-labelledby="dataSortDialogTitle" data-data-dialog="sort">
      <div class="excel-mini-titlebar" id="dataSortDialogTitle">並べ替え</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="sort-dialog-toolbar">
          <button class="excel-mini-button" type="button" data-data-dialog-action="sort-add-level">レベルの追加(A)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="sort-delete-level">レベルの削除(D)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="sort-copy-level">レベルのコピー(C)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="sort-move-up">上へ移動(U)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="sort-move-down">下へ移動(O)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="sort-options-toggle">オプション...</button>
        </div>
        <div class="sort-options-panel hidden">
          <label class="data-dialog-check"><input type="checkbox" data-data-field="caseSensitive" ${caseSensitive ? "checked" : ""} />大文字と小文字を区別する</label>
          <fieldset class="data-dialog-fieldset sort-direction-fieldset">
            <legend>方向</legend>
            <label class="data-dialog-check"><input type="radio" name="sortOrientation" data-data-field="sortOrientation" data-data-dialog-action="sort-orientation" value="rows" ${orientation === "rows" ? "checked" : ""} />上から下へ</label>
            <label class="data-dialog-check"><input type="radio" name="sortOrientation" data-data-field="sortOrientation" data-data-dialog-action="sort-orientation" value="columns" ${orientation === "columns" ? "checked" : ""} />左から右へ</label>
          </fieldset>
        </div>
        <label class="data-dialog-check"><input type="checkbox" data-data-field="header" data-data-dialog-action="sort-header-toggle" ${hasHeader ? "checked" : ""} /><span data-sort-header-label>${orientation === "columns" ? "先頭列をデータの見出しとして使用する" : "先頭行をデータの見出しとして使用する"}</span></label>
        <div class="sort-level-grid" role="group" aria-label="並べ替え条件" data-sort-level-grid>
          ${sortDialogLevelGridHtml(sheet, dialogRange, orientation, hasHeader, state.dataDialog.sortLevels, state.dataDialog.sortActiveLevel)}
        </div>
        <div class="data-dialog-summary sort-dialog-summary" data-sort-dialog-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
	    </div>`;
  document.body.appendChild(overlay);
  syncSortDialog(overlay.querySelector("[data-data-dialog='sort']"), { capture: false });
  overlay.querySelector("[data-data-sort-key='0']")?.focus({ preventScroll: true });
}



function sortStateForDialog(sheet, selectedRange) {
  const sortState = sheet?.sortState;
  if (!sortState?.range || !sortState.keys?.length || !selectedRange) return null;
  const sortRange = sortState.range;
  if (
    rangesEqual(sortRange, selectedRange) ||
    rangeContains(sortRange, selectedRange) ||
    rangeContains(selectedRange, sortRange) ||
    rangesIntersect(sortRange, selectedRange)
  ) {
    return sortState;
  }
  return null;
}



function sortDialogLevelsFromSortState(sortState, orientation = "rows") {
  const axisKey = orientation === "columns" ? "row" : "col";
  return (sortState.keys || []).map((key) => ({
    axisValue: Number(key?.[axisKey]),
    sortOn: ["values", "cellColor", "fontColor"].includes(key?.sortOn) ? key.sortOn : "values",
    direction: key?.direction === "desc" ? "desc" : "asc",
    customList: sortCustomListById(key?.customList)?.id || "",
    color: normalizeSortColor(key?.color || ""),
    colorPosition: key?.colorPosition === "bottom" ? "bottom" : "top",
  }));
}



function defaultSortDialogLevel(range, orientation = "rows") {
  return {
    axisValue: orientation === "columns" ? range.top : range.left,
    sortOn: "values",
    direction: "asc",
  };
}



function sortDialogLevelGridHtml(sheet, range, orientation, hasHeader, levels, activeIndex = 0) {
  const axisCaption = orientation === "columns" ? "行" : "列";
  const normalizedLevels = normalizeSortDialogLevels(levels, range, orientation);
  return `
    <span></span><span data-sort-axis-caption>${axisCaption}</span><span>並べ替えのキー</span><span>順序</span>
    ${normalizedLevels.map((level, index) => sortDialogLevelRowHtml(sheet, range, orientation, hasHeader, level, index, activeIndex)).join("")}`;
}



function sortDialogLevelRowHtml(sheet, range, orientation, hasHeader, level, index, activeIndex = 0) {
  const active = index === activeIndex;
  const sortOn = level.sortOn || "values";
  return `
    <button class="sort-level-selector${active ? " is-active" : ""}" type="button" data-data-dialog-action="sort-level-select" data-sort-level-index="${index}" aria-label="${index + 1} 番目の条件を選択">${active ? "▶" : ""}</button>
    <div class="sort-level-row${active ? " is-active" : ""}" data-sort-level-row="${index}">
      <select data-data-sort-key="${index}" data-sort-dialog-sync="key" aria-label="${index + 1} 番目のキー">${sortDialogKeyOptions(sheet, range, orientation, hasHeader, false, level.axisValue).join("")}</select>
      <select data-data-sort-on="${index}" data-sort-dialog-sync="sort-on" aria-label="${index + 1} 番目の並べ替えのキー">
        <option value="values" ${sortOn === "values" ? "selected" : ""}>セルの値</option>
        <option value="cellColor" ${sortOn === "cellColor" ? "selected" : ""}>セルの色</option>
        <option value="fontColor" ${sortOn === "fontColor" ? "selected" : ""}>フォントの色</option>
      </select>
      ${sortDialogOrderControlHtml(sheet, range, orientation, hasHeader, level, index)}
    </div>`;
}



function sortDialogOrderControlHtml(sheet, range, orientation, hasHeader, level, index) {
  const sortOn = level.sortOn || "values";
  if (sortOn === "cellColor" || sortOn === "fontColor") {
    const colorOptions = sortDialogColorOptions(sheet, range, orientation, hasHeader, level.axisValue, sortOn, level.color);
    const position = level.colorPosition === "bottom" ? "bottom" : "top";
    return `<div class="sort-order-color-controls">
      <select data-data-sort-color="${index}" data-sort-dialog-sync="color" aria-label="${index + 1} 番目の色">${colorOptions.join("")}</select>
      <select data-data-sort-position="${index}" data-sort-dialog-sync="position" aria-label="${index + 1} 番目の色の位置">
        <option value="top" ${position === "top" ? "selected" : ""}>上</option>
        <option value="bottom" ${position === "bottom" ? "selected" : ""}>下</option>
      </select>
    </div>`;
  }
  const direction = level.direction === "desc" ? "desc" : "asc";
  return `<select data-data-sort-order="${index}" data-sort-dialog-sync="order" aria-label="${index + 1} 番目の順序">
    <option value="asc" ${direction === "asc" ? "selected" : ""}>昇順</option>
    <option value="desc" ${direction === "desc" ? "selected" : ""}>降順</option>
    ${sortDialogCustomListOptions(level.customList)}
  </select>`;
}



function sortDialogCustomListOptions(selected = "") {
  return SORT_CUSTOM_LISTS.map((list) => {
    const value = `custom:${list.id}`;
    return `<option value="${escapeAttr(value)}" ${selected === list.id ? "selected" : ""}>${escapeHtml(list.label)}</option>`;
  }).join("");
}



function sortDialogColorOptions(sheet, range, orientation, hasHeader, axisValue, sortOn, selectedColor = "") {
  const colors = sortDialogColorItems(sheet, range, orientation, hasHeader, axisValue, sortOn);
  const selected = normalizeSortColor(selectedColor) || colors[0] || "";
  const options = colors.map((color) => `<option value="${escapeAttr(color)}" ${color === selected ? "selected" : ""}>${escapeHtml(color)}</option>`);
  if (!options.length) options.push('<option value="">色なし</option>');
  return options;
}



function sortDialogColorItems(sheet, range, orientation, hasHeader, axisValue, sortOn) {
  const colors = new Set();
  if (orientation === "columns") {
    const firstCol = hasHeader ? range.left + 1 : range.left;
    for (let col = firstCol; col <= range.right; col += 1) {
      const color = sortDialogCellColor(sheet, axisValue, col, sortOn);
      if (color) colors.add(color);
    }
  } else {
    const firstRow = hasHeader ? range.top + 1 : range.top;
    for (let row = firstRow; row <= range.bottom; row += 1) {
      const color = sortDialogCellColor(sheet, row, axisValue, sortOn);
      if (color) colors.add(color);
    }
  }
  return [...colors].sort((left, right) => left.localeCompare(right));
}



function sortDialogCellColor(sheet, row, col, sortOn) {
  const cell = sheet.cells[cellKey(row, col)];
  const css = cellEffectiveCss(sheet, cell);
  if (sortOn === "fontColor") return cellHasFontColor(cell) ? normalizeSortColor(css.color) : "";
  return cellHasFill(cell) ? normalizeSortColor(css.backgroundColor) : "";
}



function sortDialogKeyOptions(sheet, range, orientation = "rows", hasHeader = false, optional = false, selectedValue = null) {
  const sheetId = getSheetId(sheet.name);
  const options = optional ? ['<option value="">(なし)</option>'] : [];
  if (orientation === "columns") {
    for (let row = range.top; row <= range.bottom; row += 1) {
      const label = hasHeader ? String(getDisplayForCell(sheet, sheetId, row, range.left) || `${row}`) : `${row}`;
      options.push(`<option value="${row}" ${Number(selectedValue) === row ? "selected" : ""}>${escapeHtml(label)}</option>`);
    }
  } else {
    for (let col = range.left; col <= range.right; col += 1) {
      const label = hasHeader ? String(getDisplayForCell(sheet, sheetId, range.top, col) || columnName(col)) : columnName(col);
      options.push(`<option value="${col}" ${Number(selectedValue) === col ? "selected" : ""}>${escapeHtml(label)}</option>`);
    }
  }
  return options;
}



function syncSortDialogOrientation(dialog) {
  const range = state.dataDialog?.range;
  const sheet = activeSheet();
  if (!dialog || !sheet || !range) return;
  captureSortDialogLevels(dialog);
  const orientation = dialog.querySelector("[data-data-field='sortOrientation']:checked")?.value || "rows";
  const hasHeader = Boolean(dialog.querySelector("[data-data-field='header']")?.checked);
  const axisCaption = dialog.querySelector("[data-sort-axis-caption]");
  if (axisCaption) axisCaption.textContent = orientation === "columns" ? "行" : "列";
  const headerLabel = dialog.querySelector("[data-sort-header-label]");
  if (headerLabel) headerLabel.textContent = orientation === "columns" ? "先頭列をデータの見出しとして使用する" : "先頭行をデータの見出しとして使用する";
  state.dataDialog.sortOrientation = orientation;
  state.dataDialog.sortLevels = normalizeSortDialogLevels(state.dataDialog.sortLevels, range, orientation);
  rerenderSortDialogLevels(dialog);
}



function captureSortDialogLevels(dialog) {
  if (!dialog || state.dataDialog?.type !== "sort") return [];
  const orientation = dialog.querySelector("[data-data-field='sortOrientation']:checked")?.value || state.dataDialog.sortOrientation || "rows";
  const levels = Array.from(dialog.querySelectorAll("[data-sort-level-row]")).map((row) => {
    const index = row.dataset.sortLevelRow;
    const sortOn = row.querySelector(`[data-data-sort-on="${index}"]`)?.value || "values";
    const level = {
      axisValue: Number(row.querySelector(`[data-data-sort-key="${index}"]`)?.value),
      sortOn,
    };
    if (sortOn === "cellColor" || sortOn === "fontColor") {
      level.color = normalizeSortColor(row.querySelector(`[data-data-sort-color="${index}"]`)?.value || "");
      level.colorPosition = row.querySelector(`[data-data-sort-position="${index}"]`)?.value === "bottom" ? "bottom" : "top";
      level.direction = "asc";
      return level;
    }
    return {
      ...level,
      ...sortDialogOrderFromValue(row.querySelector(`[data-data-sort-order="${index}"]`)?.value || "asc"),
    };
  });
  state.dataDialog.sortOrientation = orientation;
  state.dataDialog.sortLevels = levels;
  return levels;
}



function normalizeSortDialogLevels(levels, range, orientation = "rows") {
  const source = Array.isArray(levels) && levels.length ? levels : [defaultSortDialogLevel(range, orientation)];
  const min = orientation === "columns" ? range.top : range.left;
  const max = orientation === "columns" ? range.bottom : range.right;
  return source.map((level) => {
    const axisValue = Number(level?.axisValue);
    return {
      axisValue: Number.isFinite(axisValue) && axisValue >= min && axisValue <= max ? axisValue : min,
      sortOn: ["values", "cellColor", "fontColor"].includes(level?.sortOn) ? level.sortOn : "values",
      direction: level?.direction === "desc" ? "desc" : "asc",
      customList: sortCustomListById(level?.customList)?.id || "",
      color: normalizeSortColor(level?.color || ""),
      colorPosition: level?.colorPosition === "bottom" ? "bottom" : "top",
    };
  });
}



function sortDialogOrderFromValue(value) {
  const text = String(value || "asc");
  if (text.startsWith("custom:")) {
    const list = sortCustomListById(text.slice("custom:".length));
    if (list) return { direction: "asc", customList: list.id };
  }
  return { direction: text === "desc" ? "desc" : "asc", customList: "" };
}



function handleSortDialogControlChange(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='sort']");
  if (!dialog || state.dataDialog?.type !== "sort") return;
  captureSortDialogLevels(dialog);
  if (event.currentTarget.dataset.sortDialogSync === "key" || event.currentTarget.dataset.sortDialogSync === "sort-on") {
    rerenderSortDialogLevels(dialog);
    return;
  }
  syncSortDialog(dialog, { capture: false });
}



function sortDialogActiveIndex() {
  const levels = state.dataDialog?.sortLevels || [];
  const active = Number(state.dataDialog?.sortActiveLevel || 0);
  return Math.max(0, Math.min(Math.max(0, levels.length - 1), active));
}



function rerenderSortDialogLevels(dialog) {
  const sheet = activeSheet();
  const range = state.dataDialog?.range;
  const grid = dialog?.querySelector("[data-sort-level-grid]");
  if (!sheet || !range || !grid) return;
  const orientation = state.dataDialog.sortOrientation || dialog.querySelector("[data-data-field='sortOrientation']:checked")?.value || "rows";
  const hasHeader = Boolean(dialog.querySelector("[data-data-field='header']")?.checked);
  state.dataDialog.sortLevels = normalizeSortDialogLevels(state.dataDialog.sortLevels, range, orientation);
  state.dataDialog.sortActiveLevel = sortDialogActiveIndex();
  grid.innerHTML = sortDialogLevelGridHtml(sheet, range, orientation, hasHeader, state.dataDialog.sortLevels, state.dataDialog.sortActiveLevel);
  captureSortDialogLevels(dialog);
  syncSortDialog(dialog, { capture: false });
}



function setSortDialogActiveLevel(dialog, index) {
  if (state.dataDialog?.type !== "sort") return;
  captureSortDialogLevels(dialog);
  state.dataDialog.sortActiveLevel = Math.max(0, Math.min(state.dataDialog.sortLevels.length - 1, Number(index) || 0));
  rerenderSortDialogLevels(dialog);
}



function mutateSortDialogLevels(dialog, action) {
  const range = state.dataDialog?.range;
  if (!range) return;
  const levels = captureSortDialogLevels(dialog);
  const active = sortDialogActiveIndex();
  if (action === "sort-add-level") {
    levels.splice(active + 1, 0, defaultSortDialogLevel(range, state.dataDialog.sortOrientation || "rows"));
    state.dataDialog.sortActiveLevel = active + 1;
  } else if (action === "sort-copy-level") {
    levels.splice(active + 1, 0, { ...(levels[active] || defaultSortDialogLevel(range, state.dataDialog.sortOrientation || "rows")) });
    state.dataDialog.sortActiveLevel = active + 1;
  } else if (action === "sort-delete-level") {
    if (levels.length > 1) levels.splice(active, 1);
    else levels[0] = defaultSortDialogLevel(range, state.dataDialog.sortOrientation || "rows");
    state.dataDialog.sortActiveLevel = Math.max(0, Math.min(active, levels.length - 1));
  } else if (action === "sort-move-up" && active > 0) {
    [levels[active - 1], levels[active]] = [levels[active], levels[active - 1]];
    state.dataDialog.sortActiveLevel = active - 1;
  } else if (action === "sort-move-down" && active < levels.length - 1) {
    [levels[active + 1], levels[active]] = [levels[active], levels[active + 1]];
    state.dataDialog.sortActiveLevel = active + 1;
  }
  state.dataDialog.sortLevels = levels;
  rerenderSortDialogLevels(dialog);
}



function syncSortDialog(dialog, options = {}) {
  if (!dialog || state.dataDialog?.type !== "sort") return;
  if (options.capture !== false) captureSortDialogLevels(dialog);
  const levels = state.dataDialog.sortLevels || [];
  const active = sortDialogActiveIndex();
  const setDisabled = (action, disabled) => {
    const button = dialog.querySelector(`[data-data-dialog-action='${action}']`);
    if (button) button.disabled = Boolean(disabled);
  };
  setDisabled("sort-delete-level", levels.length <= 1);
  setDisabled("sort-move-up", active <= 0);
  setDisabled("sort-move-down", active >= levels.length - 1);
  const preview = sortDialogPreview(dialog);
  const summary = dialog.querySelector("[data-sort-dialog-summary]");
  if (summary) {
    summary.textContent = preview.message;
    summary.classList.toggle("is-warning", preview.warning || preview.blocked);
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = Boolean(preview.blocked);
}



function sortDialogPreview(dialog) {
  const validation = sortDialogValidation(dialog, { capture: false });
  if (!validation.valid) return { warning: true, blocked: true, message: validation.message };
  const { sheet, range, orientation, hasHeader, sortKeys, duplicateAxes } = validation;
  const directionLabel = orientation === "columns" ? "左から右" : "上から下";
  const dataCount = orientation === "columns"
    ? Math.max(0, range.right - (hasHeader ? range.left : range.left - 1))
    : Math.max(0, range.bottom - (hasHeader ? range.top : range.top - 1));
  const first = sortKeys[0];
  const firstLabel = sortDialogAxisLabel(sheet, range, orientation, hasHeader, orientation === "columns" ? first.row : first.col);
  const firstSummary = `${firstLabel} / ${sortDialogSortOnLabel(first.sortOn)} / ${sortDialogOrderLabel(first)}`;
  const duplicateText = duplicateAxes.length ? ` 同じ${orientation === "columns" ? "行" : "列"}が複数回指定されています。` : "";
  return {
    warning: Boolean(duplicateAxes.length),
    blocked: false,
    message: `${rangeToLabel(range)} を ${directionLabel}、${sortKeys.length} 条件で並べ替えます。対象: ${dataCount.toLocaleString("ja-JP")} ${orientation === "columns" ? "列" : "行"}。先頭条件: ${firstSummary}。${duplicateText}`.trim(),
  };
}



function sortDialogValidation(dialog, options = {}) {
  const sheet = activeSheet();
  const range = state.dataDialog?.range;
  if (!dialog || !sheet || !range) return { valid: false, message: "並べ替え範囲がありません。" };
  if (options.capture !== false) captureSortDialogLevels(dialog);
  const orientation = dialog.querySelector("[data-data-field='sortOrientation']:checked")?.value || state.dataDialog.sortOrientation || "rows";
  const hasHeader = Boolean(dialog.querySelector("[data-data-field='header']")?.checked);
  const levels = normalizeSortDialogLevels(state.dataDialog.sortLevels || [], range, orientation);
  const sortKeys = sortDialogKeysFromLevels(levels, orientation);
  if (!sortKeys.length) return { valid: false, message: "並べ替えのキーを指定してください。" };
  const dataStart = orientation === "columns" ? (hasHeader ? range.left + 1 : range.left) : (hasHeader ? range.top + 1 : range.top);
  const dataEnd = orientation === "columns" ? range.right : range.bottom;
  if (dataStart >= dataEnd) return { valid: false, message: "並べ替えるデータがありません。" };
  const mergeConflict = sortDialogMergeConflict(sheet, range, orientation, hasHeader);
  if (mergeConflict) return { valid: false, message: mergeConflict };
  const invalidColorKey = sortKeys.find((key) => (key.sortOn === "cellColor" || key.sortOn === "fontColor") && !key.color);
  if (invalidColorKey) return { valid: false, message: "色で並べ替えるには、色を選択してください。" };
  const axisValues = sortKeys.map((key) => (orientation === "columns" ? key.row : key.col));
  const duplicateAxes = axisValues.filter((axis, index) => axisValues.indexOf(axis) !== index);
  return { valid: true, sheet, range, orientation, hasHeader, sortKeys, duplicateAxes };
}



function sortDialogMergeConflict(sheet, range, orientation = "rows", hasHeader = false) {
  if (!sheet?.merges?.length || !range) return "";
  const dataRange = orientation === "columns"
    ? { ...range, left: hasHeader ? range.left + 1 : range.left }
    : { ...range, top: hasHeader ? range.top + 1 : range.top };
  if (dataRange.top > dataRange.bottom || dataRange.left > dataRange.right) return "";
  const merges = mergesIntersectingRange(sheet, dataRange);
  if (!merges.length) return "";
  const hasPartialMerge = merges.some((merge) => !rangeContains(dataRange, merge));
  if (hasPartialMerge) return "この操作を行うには、結合セルが並べ替え範囲内に完全に含まれている必要があります。";
  const firstHeight = merges[0].bottom - merges[0].top + 1;
  const firstWidth = merges[0].right - merges[0].left + 1;
  const sameSize = merges.every((merge) => merge.bottom - merge.top + 1 === firstHeight && merge.right - merge.left + 1 === firstWidth);
  if (!sameSize) return "この操作を行うには、すべての結合セルを同じサイズにする必要があります。";
  return "結合セルを含む範囲は並べ替えできません。結合を解除してから実行してください。";
}



function sortDialogKeysFromLevels(levels, orientation = "rows") {
  return (levels || [])
    .map((level) => {
      const axisValue = Number(level.axisValue);
      const key = {
        sortOn: level.sortOn || "values",
        direction: level.direction || "asc",
      };
      if (level.customList) key.customList = level.customList;
      if (level.sortOn === "cellColor" || level.sortOn === "fontColor") {
        key.color = level.color || "";
        key.colorPosition = level.colorPosition === "bottom" ? "bottom" : "top";
      }
      if (orientation === "columns") key.row = axisValue;
      else key.col = axisValue;
      return key;
    })
    .filter((key) => {
      const axisValue = orientation === "columns" ? key.row : key.col;
      return Number.isFinite(axisValue) && axisValue > 0;
    });
}



function sortDialogAxisLabel(sheet, range, orientation, hasHeader, axisValue) {
  const sheetId = getSheetId(sheet.name);
  if (orientation === "columns") {
    if (hasHeader) return String(getDisplayForCell(sheet, sheetId, axisValue, range.left) || `${axisValue}`);
    return `${axisValue}`;
  }
  if (hasHeader) return String(getDisplayForCell(sheet, sheetId, range.top, axisValue) || columnName(axisValue));
  return columnName(axisValue);
}



function sortDialogSortOnLabel(sortOn) {
  if (sortOn === "cellColor") return "セルの色";
  if (sortOn === "fontColor") return "フォントの色";
  return "セルの値";
}



function sortDialogOrderLabel(key) {
  if (key.sortOn === "cellColor" || key.sortOn === "fontColor") {
    return `${key.color || "色なし"} を${key.colorPosition === "bottom" ? "下" : "上"}に表示`;
  }
  const customList = sortCustomListById(key.customList);
  if (customList) return customList.label.replace(/^ユーザー設定リスト: /, "");
  return key.direction === "desc" ? "降順" : "昇順";
}



function sortRangeRows(range, keyCol, direction = "asc", hasHeader = false) {
  return sortRangeRowsByKeys(range, [{ col: keyCol, direction }], hasHeader);
}



function sortRangeRowsByKeys(range, sortKeys, hasHeader = false, options = {}) {
  if (options.orientation === "columns") {
    return sortRangeColumnsByKeys(range, sortKeys, hasHeader, options);
  }
  const sheet = options.sheet || activeSheet();
  if (options.commitHistory !== false && !ensureSheetProtectionPermission(sheet, "allowSort", "保護されたシートでは並べ替えを使用できません。")) return false;
  const firstRow = hasHeader ? range.top + 1 : range.top;
  if (firstRow >= range.bottom) {
    if (!options.quiet) showMessageBox("並べ替えるデータがありません。", { title: "Cht WebSheet" });
    return false;
  }
  const activeKeys = (sortKeys || []).filter((item) => item?.col >= range.left && item.col <= range.right);
  if (!activeKeys.length) {
    if (!options.quiet) showMessageBox("並べ替えのキーを指定してください。", { title: "Cht WebSheet" });
    return false;
  }
  if (!options.skipMergeCheck) {
    const mergeConflict = sortDialogMergeConflict(sheet, range, "rows", hasHeader);
    if (mergeConflict) {
      if (!options.quiet) showMessageBox(mergeConflict, { title: "Cht WebSheet" });
      return false;
    }
  }
  const sheetBefore = options.commitHistory === false ? null : captureSheetForHistory(sheet);
  const rows = [];
  for (let row = firstRow; row <= range.bottom; row += 1) {
    const cells = [];
    for (let col = range.left; col <= range.right; col += 1) {
      cells.push(cloneCellModel(sheet.cells[cellKey(row, col)]));
    }
    rows.push({
      row,
      cells,
      rowModel: cloneRowModelForMove(sheet, row),
      values: activeKeys.map((key) => ({
        direction: key.direction === "desc" ? "desc" : "asc",
        value: sortComparableValue(sheet, row, key.col, key),
      })),
    });
  }
  rows.sort((a, b) => {
    for (let index = 0; index < activeKeys.length; index += 1) {
      const blankResult = compareSortBlankValues(a.values[index].value, b.values[index].value);
      if (blankResult != null) return blankResult;
      const result = compareSortValues(a.values[index].value, b.values[index].value, options);
      if (result !== 0) return result * (a.values[index].direction === "desc" ? -1 : 1);
    }
    return a.row - b.row;
  });
  rows.forEach((source, offset) => {
    const targetRow = firstRow + offset;
    sheet.rows[targetRow - 1] = cloneRowModelForMove(sheet, null, source.rowModel);
    source.cells.forEach((cell, colOffset) => {
      writeClonedCellContent(sheet, targetRow, range.left + colOffset, cell);
    });
  });
  if (options.recordSortState !== false) {
    sheet.sortState = normalizedSortStateForRows(range, activeKeys, hasHeader, options);
    syncSortStateToMatchingTable(sheet);
  }
  if (options.commitHistory === false) return true;
  commitDataOperation(sheet, sheetBefore, range, `${rangeToLabel(range)} 並べ替え`, `${rangeToLabel(range)} を並べ替えました。`);
  return true;
}



function sortRangeColumnsByKeys(range, sortKeys, hasHeader = false, options = {}) {
  const sheet = options.sheet || activeSheet();
  if (options.commitHistory !== false && !ensureSheetProtectionPermission(sheet, "allowSort", "保護されたシートでは並べ替えを使用できません。")) return false;
  const firstCol = hasHeader ? range.left + 1 : range.left;
  if (firstCol >= range.right) {
    if (!options.quiet) showMessageBox("並べ替えるデータがありません。", { title: "Cht WebSheet" });
    return false;
  }
  const activeKeys = (sortKeys || []).filter((item) => item?.row >= range.top && item.row <= range.bottom);
  if (!activeKeys.length) {
    if (!options.quiet) showMessageBox("並べ替えのキーを指定してください。", { title: "Cht WebSheet" });
    return false;
  }
  if (!options.skipMergeCheck) {
    const mergeConflict = sortDialogMergeConflict(sheet, range, "columns", hasHeader);
    if (mergeConflict) {
      if (!options.quiet) showMessageBox(mergeConflict, { title: "Cht WebSheet" });
      return false;
    }
  }
  const sheetBefore = options.commitHistory === false ? null : captureSheetForHistory(sheet);
  const columns = [];
  for (let col = firstCol; col <= range.right; col += 1) {
    const cells = [];
    for (let row = range.top; row <= range.bottom; row += 1) {
      cells.push(cloneCellModel(sheet.cells[cellKey(row, col)]));
    }
    columns.push({
      col,
      cells,
      columnModel: cloneColumnModelForMove(sheet, col),
      values: activeKeys.map((key) => ({
        direction: key.direction === "desc" ? "desc" : "asc",
        value: sortComparableValue(sheet, key.row, col, key),
      })),
    });
  }
  columns.sort((a, b) => {
    for (let index = 0; index < activeKeys.length; index += 1) {
      const blankResult = compareSortBlankValues(a.values[index].value, b.values[index].value);
      if (blankResult != null) return blankResult;
      const result = compareSortValues(a.values[index].value, b.values[index].value, options);
      if (result !== 0) return result * (a.values[index].direction === "desc" ? -1 : 1);
    }
    return a.col - b.col;
  });
  columns.forEach((source, offset) => {
    const targetCol = firstCol + offset;
    sheet.columns[targetCol - 1] = cloneColumnModelForMove(sheet, null, source.columnModel);
    source.cells.forEach((cell, rowOffset) => {
      writeClonedCellContent(sheet, range.top + rowOffset, targetCol, cell);
    });
  });
  if (options.recordSortState !== false) {
    sheet.sortState = normalizedSortStateForColumns(range, activeKeys, hasHeader, options);
    syncSortStateToMatchingTable(sheet);
  }
  if (options.commitHistory === false) return true;
  commitDataOperation(sheet, sheetBefore, range, `${rangeToLabel(range)} 並べ替え`, `${rangeToLabel(range)} を左から右へ並べ替えました。`);
  return true;
}



function cloneRowModelForMove(sheet, row, sourceModel = null) {
  const model = sourceModel || (Number.isFinite(Number(row)) ? sheet?.rows?.[Number(row) - 1] : null);
  return { ...defaultRowModel(sheet), ...(model || {}) };
}



function cloneColumnModelForMove(sheet, col, sourceModel = null) {
  const model = sourceModel || (Number.isFinite(Number(col)) ? sheet?.columns?.[Number(col) - 1] : null);
  return { ...defaultColumnModel(sheet), ...(model || {}) };
}



function normalizedSortStateForRows(range, sortKeys, hasHeader = false, options = {}) {
  return {
    orientation: "rows",
    range: { ...range },
    hasHeader: Boolean(hasHeader),
    caseSensitive: Boolean(options.caseSensitive),
    keys: normalizeSortStateKeys(sortKeys, "col"),
  };
}



function normalizedSortStateForColumns(range, sortKeys, hasHeader = false, options = {}) {
  return {
    orientation: "columns",
    range: { ...range },
    hasHeader: Boolean(hasHeader),
    caseSensitive: Boolean(options.caseSensitive),
    keys: normalizeSortStateKeys(sortKeys, "row"),
  };
}



function syncSortStateToMatchingTable(sheet) {
  const sortState = sheet?.sortState;
  if (!sheet?.tables?.length || !sortState?.range) return;
  const table = tableForAutoFilterRange(sheet, sortState.range);
  if (!table) return;
  const cloned = cloneJsonLikeValue(sortState);
  table.sortState = cloned;
  if (table.autoFilterState?.range) {
    table.autoFilterState.sortState = cloneJsonLikeValue(sortState);
  }
}



function clearSortStateFromMatchingTable(sheet, range = null) {
  if (!sheet?.tables?.length) return;
  (sheet.tables || []).forEach((table) => {
    const tableRange = tableAutoFilterRange(table);
    const sortRange = range || table.sortState?.range || table.autoFilterState?.sortState?.range || null;
    if (!tableRange || !sortRange || !rangeRectEquals(tableRange, sortRange)) return;
    delete table.sortState;
    if (table.autoFilterState) delete table.autoFilterState.sortState;
  });
}



function normalizeSortStateKeys(sortKeys, axisKey) {
  return (sortKeys || [])
    .map((key) => ({
      [axisKey]: Number(key?.[axisKey]),
      direction: key?.direction === "desc" ? "desc" : "asc",
      ...(key?.sortOn ? { sortOn: key.sortOn } : {}),
      ...(key?.customList ? { customList: key.customList } : {}),
      ...(key?.color ? { color: key.color } : {}),
      ...(key?.colorPosition ? { colorPosition: key.colorPosition } : {}),
    }))
    .filter((key) => Number.isFinite(key[axisKey]));
}



function sortComparableValue(sheet, row, col, sortKey = "values") {
  const sortOn = typeof sortKey === "string" ? sortKey : sortKey?.sortOn || "values";
  const cell = sheet.cells[cellKey(row, col)];
  if (sortOn === "cellColor" || sortOn === "fontColor") {
    return sortComparableColorValue(sheet, cell, sortOn === "cellColor" ? "backgroundColor" : "color", sortKey);
  }
  const value = cell?.formula ? getDisplayForCell(sheet, getSheetId(sheet.name), row, col) : cell?.raw ?? cell?.cached ?? cellDisplayText(cell);
  if (value == null || String(value).trim() === "") return { type: "blank", value: "" };
  if (typeof value === "number") return { type: "number", value };
  const rawDate = comparableDateSerial(value);
  if (rawDate != null) return { type: "date", value: rawDate };
  const text = String(value ?? "");
  const numeric = Number(text.replace(/,/g, ""));
  if (text.trim() && Number.isFinite(numeric)) return { type: "number", value: numeric };
  const date = comparableDateSerial(text);
  if (date != null) return { type: "date", value: date };
  const customRank = sortCustomListRank(sortKey?.customList, text);
  if (customRank != null) return { type: "custom", value: customRank, text };
  return { type: "text", value: text };
}



function comparableDateSerial(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Math.floor(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()) / 86400000);
  }
  return autoFilterDateSerial(value);
}



function sortCustomListById(id) {
  return SORT_CUSTOM_LISTS.find((list) => list.id === id) || null;
}



function sortCustomListRank(listId, value) {
  const list = sortCustomListById(listId);
  if (!list) return null;
  const normalized = normalizeCustomSortText(value);
  for (let index = 0; index < list.values.length; index += 1) {
    const candidates = [list.values[index], ...(list.aliases?.[index] || [])].map(normalizeCustomSortText);
    if (candidates.includes(normalized)) return index;
  }
  return null;
}



function normalizeCustomSortText(value) {
  return String(value ?? "").trim().toLocaleLowerCase("ja-JP");
}



function sortComparableColorValue(sheet, cell, cssProperty, sortKey = {}) {
  const css = cellEffectiveCss(sheet, cell);
  const color =
    cssProperty === "backgroundColor" && !cellHasFill(cell)
      ? ""
      : cssProperty === "color" && !cellHasFontColor(cell)
        ? ""
        : normalizeSortColor(css?.[cssProperty]);
  const requested = normalizeSortColor(sortKey?.color || "");
  if (requested) {
    const match = color === requested;
    const top = sortKey?.colorPosition !== "bottom";
    return { type: "rank", value: match ? (top ? 0 : 1) : (top ? 1 : 0) };
  }
  return {
    type: "text",
    value: color || "\uffff",
  };
}



function normalizeSortColor(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text || text === "transparent" || text === "rgba(0, 0, 0, 0)") return "";
  const hex = text.replace(/^#/, "");
  if (/^[0-9a-f]{8}$/.test(hex)) return `#${hex.slice(2)}`;
  if (/^[0-9a-f]{6}$/.test(hex)) return `#${hex}`;
  return text;
}



function compareSortValues(a, b, options = {}) {
  if (a.type === "rank" || b.type === "rank") return (Number(a.value) || 0) - (Number(b.value) || 0);
  if (a.type === "custom" || b.type === "custom") {
    if (a.type === "custom" && b.type === "custom") return a.value - b.value;
    return a.type === "custom" ? -1 : 1;
  }
  if (a.type === "date" || b.type === "date") {
    if (a.type === "date" && b.type === "date") return a.value - b.value;
    return a.type === "date" ? -1 : 1;
  }
  if (a.type === "number" && b.type === "number") return a.value - b.value;
  if (a.type === "number") return -1;
  if (b.type === "number") return 1;
  return a.value.localeCompare(b.value, "ja", {
    numeric: true,
    sensitivity: options.caseSensitive ? "variant" : "base",
  });
}



function compareSortBlankValues(a, b) {
  const aBlank = a?.type === "blank";
  const bBlank = b?.type === "blank";
  if (aBlank === bBlank) return null;
  return aBlank ? 1 : -1;
}



function inferHeaderRow(sheet, range) {
  if (range.bottom <= range.top) return false;
  let firstText = 0;
  let secondNumeric = 0;
  for (let col = range.left; col <= range.right; col += 1) {
    const first = sheet.cells[cellKey(range.top, col)];
    const second = sheet.cells[cellKey(range.top + 1, col)];
    if (String(cellDisplayText(first)).trim() && !Number.isFinite(Number(first?.raw))) firstText += 1;
    if (Number.isFinite(Number(second?.raw ?? cellDisplayText(second)))) secondNumeric += 1;
  }
  return firstText > 0 && (secondNumeric > 0 || range.right > range.left);
}



function toggleAutoFilterForSelection() {
  const context = selectedDataRange("フィルター", { expandCurrentRegion: true });
  if (!context) return;
  const sheet = context.sheet;
  if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではオートフィルターを変更できません。")) return;
  const range = autoFilterRangeFromSelection(sheet, context.range);
  const current = sheet.autoFilter?.range;
  const remove = current && shouldRemoveAutoFilterForSelection(current, range, context.range);
  if (toggleAutoFilterForSheetGroup(context, remove)) return;
  const sheetBefore = captureSheetForHistory(sheet);
  if (remove) {
    removeAutoFilter(sheet);
    commitDataOperation(sheet, sheetBefore, current, `${rangeToLabel(current)} フィルター解除`, "フィルターを解除しました。");
    return;
  }
  sheet.autoFilter = { range: { ...range }, criteria: {} };
  syncAutoFilterStateToMatchingTable(sheet);
  reapplyAutoFilter({ silent: true });
  commitDataOperation(sheet, sheetBefore, range, `${rangeToLabel(range)} フィルター`, autoFilterStatusMessage(`${rangeToLabel(range)} にフィルターを設定しました`, sheet, range));
}



function toggleAutoFilterForSheetGroup(context, remove) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !context?.range) return false;
  const activePoint = state.selected?.sheetIndex === state.activeSheetIndex
    ? { row: state.selected.row, col: state.selected.col }
    : { row: context.range.top, col: context.range.left };
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではオートフィルターを変更できません。")) return true;
    const selectionRange = materializeSelectionRangeForSheet(sheet, context.range, sheetIndex);
    if (!selectionRange) continue;
    const range = autoFilterRangeFromSelection(sheet, selectionRange, {
      activePoint: { sheetIndex, row: activePoint.row, col: activePoint.col },
      sheetIndex,
    });
    items.push({ sheetIndex, sheet, selectionRange, range });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      if (remove) {
        removeAutoFilter(item.sheet);
        return;
      }
      item.sheet.autoFilter = { range: { ...item.range }, criteria: {} };
      syncAutoFilterStateToMatchingTable(item.sheet);
      reapplyAutoFilter({ silent: true, sheet: item.sheet, sheetIndex: item.sheetIndex });
    });
    state.hf = buildFormulaEngine(state.model);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: context.range,
      selectionAfter: context.range,
      selectionBeforeRanges: activeSelectionRanges().map((range) => ({ ...range })),
      selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
      label: `${sheetIndexes.length} シート ${remove ? "フィルター解除" : "フィルター"}`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(remove
      ? `${sheetIndexes.length} 個のシートのフィルターを解除しました。`
      : `${sheetIndexes.length} 個のシートにフィルターを設定しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`フィルターを変更できませんでした：${error.message || error}`);
  }
  return true;
}



function autoFilterRangeFromSelection(sheet, range, options = {}) {
  const tableRange = autoFilterTableRangeFromSelection(sheet, range, options);
  if (tableRange) return tableRange;
  const explicitHeaderSelection = !range.fullRow && range.top === range.bottom && range.left !== range.right;
  const headerRowRegion = autoFilterHeaderRowRegionFromSelection(sheet, range, options.sheetIndex ?? state.activeSheetIndex, {
    expandSideways: !explicitHeaderSelection,
  });
  if (headerRowRegion) return headerRowRegion;
  if (explicitHeaderSelection) return range;
  if (range.top !== range.bottom || range.left !== range.right) return range;
  const point = options.activePoint || (state.selected?.sheetIndex === state.activeSheetIndex
    ? state.selected
    : { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left });
  const region = currentRegionFromPoint(sheet, point, options.sheetIndex ?? point.sheetIndex ?? state.activeSheetIndex);
  if (region) return region;
  return range;
}



function shouldRemoveAutoFilterForSelection(current, resolvedRange, selectedRange, options = {}) {
  if (!current) return false;
  if (rangesEqual(current, resolvedRange)) return true;
  const active = options.activePoint || (state.selected?.sheetIndex === state.activeSheetIndex ? state.selected : null);
  if (active && isCellInRange(current, active.row, active.col)) return true;
  return Boolean(selectedRange && rangeContains(current, selectedRange));
}



function autoFilterTableRangeFromSelection(sheet, range, options = {}) {
  if (!sheet?.tables?.length || !range) return null;
  const active = options.activePoint || (state.selected?.sheetIndex === state.activeSheetIndex ? state.selected : null);
  const ranked = (sheet.tables || [])
    .map((table) => {
      const filterRange = tableAutoFilterRange(table);
      if (!filterRange || table.showHeaderRow === false) return null;
      const containsActive = Boolean(active && isCellInRange(filterRange, active.row, active.col));
      const containsSelection = rangeContains(filterRange, range);
      const intersectsSelection = rangesIntersect(filterRange, range);
      if (!containsActive && !containsSelection && !intersectsSelection) return null;
      const area = (filterRange.bottom - filterRange.top + 1) * (filterRange.right - filterRange.left + 1);
      return { range: filterRange, containsActive, containsSelection, intersectsSelection, area };
    })
    .filter(Boolean)
    .sort((a, b) =>
      Number(b.containsActive) - Number(a.containsActive) ||
      Number(b.containsSelection) - Number(a.containsSelection) ||
      Number(b.intersectsSelection) - Number(a.intersectsSelection) ||
      a.area - b.area ||
      a.range.top - b.range.top ||
      a.range.left - b.range.left,
    );
  return ranked[0]?.range || null;
}



function tableForAutoFilterRange(sheet, range) {
  if (!sheet?.tables?.length || !range) return null;
  return (sheet.tables || []).find((table) => {
    const tableRange = tableAutoFilterRange(table);
    return tableRange && rangesEqual(tableRange, range);
  }) || null;
}



function tableForAutoFilterHeaderCell(sheet, row, col) {
  if (!sheet?.tables?.length) return null;
  return (sheet.tables || []).find((table) => {
    const range = tableAutoFilterRange(table);
    return range && table.showHeaderRow !== false && table.showFilterButton !== false && row === range.top && col >= range.left && col <= range.right;
  }) || null;
}



function autoFilterStateForHeaderCell(sheet, row, col) {
  const range = sheet?.autoFilter?.range;
  if (range && row === range.top && col >= range.left && col <= range.right) return sheet.autoFilter;
  const table = tableForAutoFilterHeaderCell(sheet, row, col);
  if (!table) return null;
  const tableRange = tableAutoFilterRange(table);
  return table.autoFilterState || { range: tableRange, criteria: {}, hiddenButtons: {} };
}



function autoFilterSortKeyForHeaderCell(sheet, row, col) {
  const sheetKey = sortKeyForHeaderCell(sheet?.sortState, row, col);
  if (sheetKey) return sheetKey;
  const table = tableForAutoFilterHeaderCell(sheet, row, col);
  const tableKey = sortKeyForHeaderCell(table?.autoFilterState?.sortState || table?.sortState, row, col);
  return tableKey || null;
}



function sortKeyForHeaderCell(sortState, row, col) {
  if (!sortState?.range || !Array.isArray(sortState.keys) || !sortState.keys.length) return null;
  const range = sortState.range;
  if (sortState.orientation === "columns") {
    if (col !== range.left || row < range.top || row > range.bottom) return null;
    return sortState.keys.find((key) => Number(key?.row) === row) || null;
  }
  if (row !== range.top || col < range.left || col > range.right) return null;
  return sortState.keys.find((key) => Number(key?.col) === col) || null;
}



function autoFilterTableForSelection(sheet, range = activeSelectionRange()) {
  if (!sheet?.tables?.length || !range) return null;
  const active = state.selected?.sheetIndex === state.activeSheetIndex ? state.selected : null;
  const ranked = (sheet.tables || [])
    .map((table) => {
      const filterRange = tableAutoFilterRange(table);
      if (!filterRange || table.showHeaderRow === false || table.showFilterButton === false) return null;
      const containsActive = Boolean(active && isCellInRange(filterRange, active.row, active.col));
      const containsSelection = rangeContains(filterRange, range);
      const intersectsSelection = rangesIntersect(filterRange, range);
      if (!containsActive && !containsSelection && !intersectsSelection) return null;
      const area = (filterRange.bottom - filterRange.top + 1) * (filterRange.right - filterRange.left + 1);
      return { table, range: filterRange, containsActive, containsSelection, intersectsSelection, area };
    })
    .filter(Boolean)
    .sort((a, b) =>
      Number(b.containsActive) - Number(a.containsActive) ||
      Number(b.containsSelection) - Number(a.containsSelection) ||
      Number(b.intersectsSelection) - Number(a.intersectsSelection) ||
      a.area - b.area ||
      a.range.top - b.range.top ||
      a.range.left - b.range.left,
    );
  return ranked[0]?.table || null;
}



function activateAutoFilterTable(sheet, table) {
  const tableRange = tableAutoFilterRange(table);
  if (!sheet || !table || !tableRange) return null;
  const source = table.autoFilterState || { range: tableRange, criteria: {}, hiddenButtons: {} };
  const sourceSortState = source.sortState || table.sortState || null;
  table.showFilterButton = true;
  table.autoFilterState = {
    range: { ...tableRange },
    criteria: cloneJsonLikeValue(source.criteria || {}),
    ...(source.hiddenButtons ? { hiddenButtons: cloneJsonLikeValue(source.hiddenButtons) } : {}),
    ...(source.advancedCriteria ? { advancedCriteria: cloneJsonLikeValue(source.advancedCriteria) } : {}),
    ...(sourceSortState ? { sortState: cloneJsonLikeValue(sourceSortState) } : {}),
  };
  sheet.autoFilter = {
    range: { ...tableRange },
    criteria: cloneJsonLikeValue(table.autoFilterState.criteria || {}),
    ...(table.autoFilterState.hiddenButtons ? { hiddenButtons: cloneJsonLikeValue(table.autoFilterState.hiddenButtons) } : {}),
    ...(table.autoFilterState.advancedCriteria ? { advancedCriteria: cloneJsonLikeValue(table.autoFilterState.advancedCriteria) } : {}),
  };
  if (sourceSortState) sheet.sortState = cloneJsonLikeValue(sourceSortState);
  return sheet.autoFilter;
}



function autoFilterContextForSelection(sheet, range = activeSelectionRange()) {
  const table = autoFilterTableForSelection(sheet, range);
  if (table) {
    const filter = activateAutoFilterTable(sheet, table);
    if (filter?.range) return { type: "table", table, filter, range: filter.range };
  }
  if (sheet?.autoFilter?.range) return { type: "sheet", table: null, filter: sheet.autoFilter, range: sheet.autoFilter.range };
  return null;
}



function activateAutoFilterForHeaderCell(sheet, row, col) {
  const range = sheet?.autoFilter?.range;
  if (range && row === range.top && col >= range.left && col <= range.right) return true;
  const table = tableForAutoFilterHeaderCell(sheet, row, col);
  if (!table) return false;
  return Boolean(activateAutoFilterTable(sheet, table));
}



function syncAutoFilterStateToMatchingTable(sheet, options = {}) {
  const range = sheet?.autoFilter?.range;
  if (!sheet?.tables?.length || !range) return;
  const table = tableForAutoFilterRange(sheet, range);
  if (!table) return;
  table.showFilterButton = options.showFilterButton === false ? false : true;
  table.autoFilterState = table.showFilterButton
    ? {
        range: { ...range },
        criteria: cloneJsonLikeValue(sheet.autoFilter.criteria || {}),
        ...(sheet.autoFilter.hiddenButtons ? { hiddenButtons: cloneJsonLikeValue(sheet.autoFilter.hiddenButtons) } : {}),
        ...(sheet.autoFilter.advancedCriteria ? { advancedCriteria: cloneJsonLikeValue(sheet.autoFilter.advancedCriteria) } : {}),
        ...(sheet.sortState?.range && rangeRectEquals(sheet.sortState.range, range) ? { sortState: cloneJsonLikeValue(sheet.sortState) } : {}),
      }
    : null;
  if (table.showFilterButton && sheet.sortState?.range && rangeRectEquals(sheet.sortState.range, range)) {
    table.sortState = cloneJsonLikeValue(sheet.sortState);
  } else if (!table.showFilterButton) {
    delete table.sortState;
  }
}



function autoFilterHeaderRowRegionFromSelection(sheet, range, sheetIndex = state.activeSheetIndex, options = {}) {
  if (!sheet || !range || range.top !== range.bottom || isSingleCellRange(range)) return null;
  const used = sheetUsedRange(sheet);
  const row = range.top;
  if (!used || row < used.top || row > used.bottom) return null;
  const mergeMaps = buildMergeMaps(sheet.merges || []);
  const selectedLeft = range.fullRow ? used.left : Math.max(range.left, used.left);
  const selectedRight = range.fullRow ? used.right : Math.min(range.right, used.right);
  if (selectedLeft > selectedRight) return null;
  const probeBottom = autoFilterHeaderProbeBottom(sheet, mergeMaps, row, used);
  const span = autoFilterHeaderContinuousSpan(sheet, mergeMaps, row, selectedLeft, selectedRight, used, probeBottom, range.fullRow, options.expandSideways !== false);
  if (!span) return null;
  const left = span.left;
  const right = span.right;
  const bodyBottom = autoFilterHeaderRowBodyBottom(sheet, mergeMaps, row, left, right);
  const headerIntervals = new Map();
  for (let col = left; col <= right; col += 1) {
    const span = autoFilterHeaderSpanAt(sheet, mergeMaps, row, col, bodyBottom);
    if (!span) continue;
    const interval = {
      left: Math.max(span.left, used.left),
      right: Math.min(span.right, used.right),
    };
    if (interval.left > interval.right) continue;
    headerIntervals.set(`${interval.left}:${interval.right}`, interval);
  }
  const headerBlocks = mergedAutoFilterHeaderBlocks([...headerIntervals.values()]);
  if (!headerBlocks.length) return null;

  const selectedSpan = { ...range, left, right };
  const active = state.selected?.sheetIndex === state.activeSheetIndex ? state.selected : null;
  const ranked = headerBlocks
    .map((block) => {
      const candidate = {
        sheetIndex,
        top: row,
        left: block.left,
        right: block.right,
        bottom: autoFilterHeaderBlockBottom(sheet, mergeMaps, row, block.left, block.right, bodyBottom),
      };
      return { candidate };
    })
    .map((item) => ({
      ...item,
      containsSelection: !range.fullRow && rangeContains(item.candidate, selectedSpan),
      intersectsSelection: rangesIntersect(item.candidate, selectedSpan),
      containsActive: Boolean(active && active.row === row && active.col >= item.candidate.left && active.col <= item.candidate.right),
      area: (item.candidate.bottom - item.candidate.top + 1) * (item.candidate.right - item.candidate.left + 1),
      width: item.candidate.right - item.candidate.left + 1,
      height: item.candidate.bottom - item.candidate.top + 1,
    }))
    .filter((item) => item.intersectsSelection)
    .sort((a, b) =>
      Number(b.containsActive) - Number(a.containsActive) ||
      Number(b.containsSelection) - Number(a.containsSelection) ||
      b.area - a.area ||
      b.width - a.width ||
      b.height - a.height ||
      a.candidate.left - b.candidate.left,
    );
  const intersecting = ranked.filter((item) => item.intersectsSelection);
  if (!intersecting.length) return null;
  if (!range.fullRow && range.left !== range.right) {
    const unionLeft = Math.min(...intersecting.map((item) => item.candidate.left));
    const unionRight = Math.max(...intersecting.map((item) => item.candidate.right));
    const unionBottom = autoFilterHeaderBlockBottom(sheet, mergeMaps, row, unionLeft, unionRight, bodyBottom);
    return {
      sheetIndex,
      top: row,
      left: unionLeft,
      right: unionRight,
      bottom: Math.max(unionBottom, ...intersecting.map((item) => item.candidate.bottom)),
    };
  }
  return ranked[0]?.candidate || null;
}



function autoFilterHeaderProbeBottom(sheet, mergeMaps, headerRow, used) {
  if (!used || used.bottom <= headerRow) return headerRow;
  const scanBottom = Math.min(used.bottom, headerRow + AUTO_FILTER_HEADER_SCAN_LIMIT);
  let bottom = headerRow;
  for (let row = headerRow + 1; row <= scanBottom; row += 1) {
    if (!autoFilterBlockRowHasData(sheet, mergeMaps, row, used.left, used.right)) break;
    bottom = row;
  }
  return bottom;
}



function autoFilterHeaderContinuousSpan(sheet, mergeMaps, headerRow, selectedLeft, selectedRight, used, bodyBottom, fullRow = false, expandSideways = true) {
  const seed = {
    left: Math.max(used.left, selectedLeft),
    right: Math.min(used.right, selectedRight),
  };
  if (seed.left > seed.right) return null;
  let left = seed.left;
  let right = seed.right;

  const seedHasSignal = autoFilterHeaderSpanHasSignal(sheet, mergeMaps, headerRow, seed.left, seed.right, bodyBottom);
  if (!seedHasSignal && !fullRow) return null;

  if (expandSideways) {
    for (let col = left - 1; col >= used.left; col -= 1) {
      if (!autoFilterHeaderColumnHasSignal(sheet, mergeMaps, headerRow, col, bodyBottom)) break;
      const span = autoFilterHeaderSpanAt(sheet, mergeMaps, headerRow, col, bodyBottom);
      left = Math.max(used.left, Math.min(col, span?.left || col));
    }
    for (let col = right + 1; col <= used.right; col += 1) {
      if (!autoFilterHeaderColumnHasSignal(sheet, mergeMaps, headerRow, col, bodyBottom)) break;
      const span = autoFilterHeaderSpanAt(sheet, mergeMaps, headerRow, col, bodyBottom);
      right = Math.min(used.right, Math.max(col, span?.right || col));
    }
  }

  if (fullRow) {
    while (left <= right && !autoFilterHeaderColumnHasSignal(sheet, mergeMaps, headerRow, left, bodyBottom)) left += 1;
    while (right >= left && !autoFilterHeaderColumnHasSignal(sheet, mergeMaps, headerRow, right, bodyBottom)) right -= 1;
  }
  return left <= right ? { left, right } : null;
}



function autoFilterHeaderSpanHasSignal(sheet, mergeMaps, headerRow, left, right, bodyBottom) {
  for (let col = left; col <= right; col += 1) {
    if (autoFilterHeaderColumnHasSignal(sheet, mergeMaps, headerRow, col, bodyBottom)) return true;
  }
  return false;
}



function autoFilterHeaderColumnHasSignal(sheet, mergeMaps, headerRow, col, bodyBottom) {
  return Boolean(autoFilterHeaderSpanAt(sheet, mergeMaps, headerRow, col, bodyBottom));
}



function autoFilterHeaderSpanAt(sheet, mergeMaps, row, col, bodyBottom = row) {
  const key = cellKey(row, col);
  const masterKey = mergeMaps.coveredToMaster.get(key) || key;
  const [masterRow, masterCol] = parseCellKey(masterKey);
  const hasHeaderValue = cellHasDataForRegion(sheet, masterRow, masterCol);
  const merge = mergeMaps.master.get(masterKey);
  if (hasHeaderValue && merge && row >= merge.top && row <= merge.bottom) {
    return { left: merge.left, right: merge.right };
  }
  if (!hasHeaderValue && !autoFilterColumnHasBodyData(sheet, mergeMaps, row + 1, bodyBottom, col)) return null;
  return { left: col, right: col };
}



function mergedAutoFilterHeaderBlocks(intervals) {
  return intervals
    .sort((a, b) => a.left - b.left || a.right - b.right)
    .reduce((blocks, interval) => {
      const current = blocks[blocks.length - 1];
      if (current && interval.left <= current.right + 1) {
        current.right = Math.max(current.right, interval.right);
      } else {
        blocks.push({ ...interval });
      }
      return blocks;
    }, []);
}



function autoFilterHeaderRowBodyBottom(sheet, mergeMaps, headerRow, left, right) {
  const used = sheetUsedRange(sheet);
  if (!used) return headerRow;
  let bottom = headerRow;
  for (let row = headerRow + 1; row <= used.bottom; row += 1) {
    if (!autoFilterBlockRowHasData(sheet, mergeMaps, row, left, right)) break;
    bottom = row;
  }
  return bottom;
}



function autoFilterHeaderBlockBottom(sheet, mergeMaps, headerRow, left, right, bodyBottom = null) {
  const used = sheetUsedRange(sheet);
  const limit = bodyBottom == null ? used?.bottom : bodyBottom;
  if (!used || !limit || limit <= headerRow) return headerRow;
  let bottom = headerRow;
  for (let row = headerRow + 1; row <= Math.min(limit, used.bottom); row += 1) {
    if (!autoFilterBlockRowHasData(sheet, mergeMaps, row, left, right)) break;
    bottom = row;
  }
  return bottom;
}



function autoFilterColumnHasBodyData(sheet, mergeMaps, top, bottom, col) {
  if (bottom < top) return false;
  const seenMasters = new Set();
  for (let row = top; row <= bottom; row += 1) {
    const key = cellKey(row, col);
    const masterKey = mergeMaps.coveredToMaster.get(key) || key;
    if (seenMasters.has(masterKey)) continue;
    seenMasters.add(masterKey);
    const [masterRow, masterCol] = parseCellKey(masterKey);
    if (cellHasDataForRegion(sheet, masterRow, masterCol)) return true;
  }
  return false;
}



function autoFilterBlockRowHasData(sheet, mergeMaps, row, left, right) {
  const seenMasters = new Set();
  for (let col = left; col <= right; col += 1) {
    const key = cellKey(row, col);
    const masterKey = mergeMaps.coveredToMaster.get(key) || key;
    if (seenMasters.has(masterKey)) continue;
    seenMasters.add(masterKey);
    const [masterRow, masterCol] = parseCellKey(masterKey);
    if (cellHasDataForRegion(sheet, masterRow, masterCol)) return true;
  }
  return false;
}



function sheetUsedRange(sheet) {
  if (!sheet) return null;
  const cells = sheet.cells || {};
  const cached = sheetUsedRangeCacheBySheet.get(sheet);
  if (cached?.cells === cells) return cached.range;
  const keys = Object.entries(cells)
    .filter(([, cell]) => String(cellDisplayText(cell)).trim() || cell?.formula)
    .map(([key]) => parseCellKey(key));
  const range = keys.length
    ? {
        top: Math.min(...keys.map(([row]) => row)),
        left: Math.min(...keys.map(([, col]) => col)),
        bottom: Math.max(...keys.map(([row]) => row)),
        right: Math.max(...keys.map(([, col]) => col)),
      }
    : null;
  sheetUsedRangeCacheBySheet.set(sheet, { cells, range });
  return range;
}



function removeAutoFilter(sheet) {
  const range = sheet.autoFilter?.range;
  if (range) {
    for (let row = range.top + 1; row <= range.bottom; row += 1) {
      if (sheet.rows[row - 1]) sheet.rows[row - 1].hidden = false;
    }
    syncAutoFilterStateToMatchingTable(sheet, { showFilterButton: false });
    clearSortStateFromMatchingTable(sheet, range);
    reapplyVisibleTableAutoFilters(sheet);
  }
  sheet.autoFilter = null;
  closeAutoFilterMenu();
}



function clearAutoFilterCriteria() {
  const sheet = activeSheet();
  if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではフィルターをクリアできません。")) return;
  if (clearAutoFilterCriteriaForSheetGroup()) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const context = autoFilterContextForSelection(sheet);
  const filter = context?.filter || null;
  const filterRange = filter?.range || null;
  const tableSortState = tableSortStateForAutoFilterContext(context) || tableSortStateForFilterRange(sheet, filterRange) || visibleTableSortStates(sheet)[0]?.sortState || null;
  const sortState = sheet?.sortState || tableSortState;
  const sortRange = sortState?.range || null;
  const shouldClearSort = Boolean(sortRange && (!context || context.type !== "table" || rangeRectEquals(sortRange, filterRange)));
  if (!filterRange && !shouldClearSort) {
    setStatus("クリアするフィルターまたは並べ替えがありません。");
    return;
  }
  const targetRange = filterRange || (shouldClearSort ? sortRange : null) || activeSelectionRange();
  const hadFilterCriteria = autoFilterStateHasCriteria(filter);
  const hadSortState = shouldClearSort;
  if (filterRange) {
    filter.criteria = {};
    delete filter.advancedCriteria;
    sheet.autoFilter = filter;
    syncAutoFilterStateToMatchingTable(sheet);
    if (context?.type === "table") reapplyVisibleTableAutoFilters(sheet, { activeRange: filterRange });
    else reapplyAutoFilter({ silent: true });
  }
  if (shouldClearSort) {
    if (sheet.sortState?.range && rangeRectEquals(sheet.sortState.range, sortRange)) delete sheet.sortState;
    clearSortStateFromMatchingTable(sheet, sortRange);
  }
  const label = hadSortState ? "フィルターと並べ替え クリア" : "フィルター クリア";
  const status =
    hadFilterCriteria && hadSortState
      ? autoFilterStatusMessage("フィルター条件と並べ替え状態をクリアしました", sheet, filterRange)
      : hadSortState
        ? "並べ替え状態をクリアしました。"
        : autoFilterStatusMessage("フィルター条件をクリアしました", sheet, filterRange);
  commitDataOperation(sheet, sheetBefore, targetRange, label, status);
}



function clearAutoFilterCriteriaForSheetGroup() {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const selectionRange = activeSelectionRange();
  const activePoint = state.selected?.sheetIndex === state.activeSheetIndex
    ? { row: state.selected.row, col: state.selected.col }
    : selectionRange
      ? { row: selectionRange.top, col: selectionRange.left }
      : { row: 1, col: 1 };
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではフィルターをクリアできません。")) return true;
    const range = selectionRange
      ? materializeSelectionRangeForSheet(sheet, selectionRange, sheetIndex)
      : fallbackSingleCellRangeForSheet(sheetIndex, activePoint.row, activePoint.col);
    const context = autoFilterContextForSelection(sheet, range);
    const filter = context?.filter || null;
    const filterRange = filter?.range || null;
    const tableSortState = tableSortStateForAutoFilterContext(context) || tableSortStateForFilterRange(sheet, filterRange) || visibleTableSortStates(sheet)[0]?.sortState || null;
    const sortState = sheet?.sortState || tableSortState;
    const sortRange = sortState?.range || null;
    const shouldClearSort = Boolean(sortRange && (!context || context.type !== "table" || rangeRectEquals(sortRange, filterRange)));
    if (!filterRange && !shouldClearSort) continue;
    items.push({ sheetIndex, sheet, context, filter, filterRange, sortRange, shouldClearSort });
  }
  if (!items.length) {
    setStatus("クリアするフィルターまたは並べ替えがありません。");
    return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    let criteriaCount = 0;
    let sortCount = 0;
    items.forEach((item) => {
      if (item.filterRange) {
        if (autoFilterStateHasCriteria(item.filter)) criteriaCount += 1;
        item.filter.criteria = {};
        delete item.filter.advancedCriteria;
        item.sheet.autoFilter = item.filter;
        syncAutoFilterStateToMatchingTable(item.sheet);
        if (item.context?.type === "table") reapplyVisibleTableAutoFilters(item.sheet, { activeRange: item.filterRange });
        else reapplyAutoFilter({ silent: true, sheet: item.sheet, sheetIndex: item.sheetIndex });
      }
      if (item.shouldClearSort) {
        clearSortStateFromMatchingTable(item.sheet, item.sortRange);
        if (item.sheet.sortState?.range && rangeRectEquals(item.sheet.sortState.range, item.sortRange)) delete item.sheet.sortState;
        sortCount += 1;
      }
    });
    state.hf = buildFormulaEngine(state.model);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: selectionRange,
      selectionAfter: selectionRange,
      selectionBeforeRanges: activeSelectionRanges().map((range) => ({ ...range })),
      selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
      label: `${sheetIndexes.length} シート フィルターと並べ替え クリア`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    const detail = criteriaCount && sortCount
      ? "フィルター条件と並べ替え状態"
      : sortCount
        ? "並べ替え状態"
        : "フィルター条件";
    setStatus(`${items.length} 個のシートの${detail}をクリアしました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`フィルターをクリアできませんでした：${error.message || error}`);
  }
  return true;
}



function reapplyAutoFilter(options = {}) {
  const sheet = options.sheet || activeSheet();
  if (!options.silent && !ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではフィルターを再適用できません。")) return;
  const sheetIndex = Number.isFinite(Number(options.sheetIndex)) ? Number(options.sheetIndex) : state.activeSheetIndex;
  const filter = sheet?.autoFilter;
  const sortState = options.includeSort ? sheet?.sortState : null;
  const hasTableAutoFilters = visibleTableAutoFilters(sheet).length > 0;
  const hasTableSortStates = options.includeSort && visibleTableSortStates(sheet).length > 0;
  if (!filter?.range && !sortState && !hasTableAutoFilters && !hasTableSortStates) {
    if (!options.silent) setStatus("再適用するフィルターがありません。");
    return;
  }
  const shouldRecordHistory = !options.silent && options.recordHistory !== false;
  const sheetBefore = shouldRecordHistory ? captureSheetForHistory(sheet) : null;
  let didSort = false;
  if (sortState) didSort = applyStoredSortState(sheet, sortState);
  if (options.includeSort && reapplyVisibleTableSorts(sheet, { skipSheetSortRange: true })) didSort = true;
  const range = filter?.range;
  if (range) {
    syncAutoFilterStateToMatchingTable(sheet);
    if (!reapplyVisibleTableAutoFilters(sheet, { activeRange: range })) {
      applyAutoFilterStateToRows(sheet, filter, { resetRows: true });
    }
  } else if (hasTableAutoFilters) {
    reapplyVisibleTableAutoFilters(sheet);
  }
  if (!options.silent) {
    state.hf = buildFormulaEngine(state.model);
    if (shouldRecordHistory) {
      pushHistory({
        sheetIndex,
        sheetBefore,
        sheetAfter: captureSheetForHistory(sheet),
        selectionBefore: state.selectionRange || null,
        selectionAfter: state.selectionRange || null,
        selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
        selectionAfterRanges: activeSelectionRanges().map((item) => ({ ...item })),
        label: "フィルターと並べ替えの再適用",
      });
    }
    renderSheet();
    setStatus(didSort && range
      ? autoFilterStatusMessage("並べ替えとフィルターを再適用しました", sheet, range)
      : didSort
        ? "並べ替えを再適用しました。"
        : autoFilterStatusMessage("フィルターを再適用しました", sheet, range));
  }
}



function reapplyAutoFilterForSheetGroup(options = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではフィルターを再適用できません。")) return true;
    const hasTableAutoFilters = visibleTableAutoFilters(sheet).length > 0;
    const hasTableSortStates = options.includeSort && visibleTableSortStates(sheet).length > 0;
    if (!sheet.autoFilter?.range && !sheet.sortState?.range && !hasTableAutoFilters && !hasTableSortStates) continue;
    items.push({ sheetIndex, sheet });
  }
  if (!items.length) {
    setStatus("再適用するフィルターがありません。");
    return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => reapplyAutoFilter({
      ...options,
      silent: true,
      sheet: item.sheet,
      sheetIndex: item.sheetIndex,
    }));
    state.hf = buildFormulaEngine(state.model);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: state.selectionRange || null,
      selectionAfter: state.selectionRange || null,
      selectionBeforeRanges: activeSelectionRanges().map((range) => ({ ...range })),
      selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
      label: `${items.length} シート フィルターと並べ替えの再適用`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートのフィルターと並べ替えを再適用しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`フィルターを再適用できませんでした：${error.message || error}`);
  }
  return true;
}



function reapplyImportedAutoFiltersOnLoad() {
  if (!state.model?.sheets?.length || !state.hf) return;
  let reapplied = false;
  state.model.sheets.forEach((sheet, sheetIndex) => {
    const hasCriteria = Boolean(sheet?.autoFilter?.range && Object.keys(sheet.autoFilter.criteria || {}).length);
    const hasSortState = Boolean(sheet?.sortState?.range && sheet.sortState?.keys?.length);
    if (hasCriteria || hasSortState) {
      reapplyAutoFilter({ sheet, sheetIndex, silent: true, includeSort: hasSortState });
      reapplied = true;
    }
    if (reapplyTableAutoFiltersOnLoad(sheet)) reapplied = true;
  });
  if (reapplied) state.hf = buildFormulaEngine(state.model);
}



function autoFilterStateHasCriteria(filter) {
  return Boolean(filter?.range && (Object.keys(filter.criteria || {}).length || filter.advancedCriteria));
}



function sortStateHasKeys(sortState) {
  return Boolean(sortState?.range && Array.isArray(sortState.keys) && sortState.keys.length);
}



function tableSortStateForAutoFilterContext(context) {
  if (context?.type !== "table" || !context.table) return null;
  const sortState = context.table.autoFilterState?.sortState || context.table.sortState || null;
  return sortStateHasKeys(sortState) ? sortState : null;
}



function tableSortStateForFilterRange(sheet, range) {
  if (!sheet || !range) return null;
  return visibleTableSortStates(sheet, { activeRange: range })[0]?.sortState || null;
}



function visibleTableSortStates(sheet, options = {}) {
  const activeRange = options.activeRange || null;
  return (sheet?.tables || [])
    .filter((table) => table?.showFilterButton !== false)
    .map((table) => {
      const sortState = table.autoFilterState?.sortState || table.sortState || null;
      const range = sortState?.range || tableAutoFilterRange(table);
      return { table, sortState, range };
    })
    .filter((item) => sortStateHasKeys(item.sortState) && (!activeRange || rangeRectEquals(item.range || item.sortState.range, activeRange)));
}



function reapplyTableAutoFiltersOnLoad(sheet) {
  const didFilter = reapplyVisibleTableAutoFilters(sheet, { criteriaOnly: true });
  const didSort = reapplyVisibleTableSorts(sheet, { skipSheetSortRange: true });
  return didFilter || didSort;
}



function visibleTableAutoFilters(sheet, options = {}) {
  const filters = (sheet?.tables || [])
    .filter((table) => table?.showFilterButton !== false && table.autoFilterState?.range)
    .map((table) => table.autoFilterState);
  if (options.criteriaOnly) return filters.filter(autoFilterStateHasCriteria);
  return filters;
}



function reapplyVisibleTableAutoFilters(sheet, options = {}) {
  const filters = visibleTableAutoFilters(sheet, options);
  if (!filters.length) return false;
  const activeRange = options.activeRange || null;
  if (activeRange && !filters.some((filter) => rangesEqual(filter.range, activeRange))) return false;
  filters.forEach((filter) => resetRowsForAutoFilterRange(sheet, filter.range));
  filters.filter(autoFilterStateHasCriteria).forEach((filter) => applyAutoFilterStateToRows(sheet, filter, { resetRows: false }));
  return true;
}



function reapplyVisibleTableSorts(sheet, options = {}) {
  const items = visibleTableSortStates(sheet, options);
  if (!items.length) return false;
  let didSort = false;
  items.forEach((item) => {
    if (
      options.skipSheetSortRange &&
      sheet?.sortState?.range &&
      rangeRectEquals(sheet.sortState.range, item.sortState.range)
    ) {
      return;
    }
    if (applyStoredSortState(sheet, item.sortState)) didSort = true;
  });
  return didSort;
}



function resetRowsForAutoFilterRange(sheet, range) {
  if (!sheet || !range) return;
  for (let row = range.top + 1; row <= range.bottom; row += 1) {
    if (sheet.rows[row - 1]) sheet.rows[row - 1].hidden = false;
  }
}



function applyAutoFilterStateToRows(sheet, filter, options = {}) {
  const range = filter?.range;
  if (!sheet || !range) return false;
  if (options.resetRows !== false) resetRowsForAutoFilterRange(sheet, range);
  const advancedVisibleRows = filter?.advancedCriteria
    ? new Set(advancedFilterMatchingRows(sheet, range, filter.advancedCriteria).map((item) => item.row))
    : null;
  if (advancedVisibleRows) {
    for (let row = range.top + 1; row <= range.bottom; row += 1) {
      if (!advancedVisibleRows.has(row) && sheet.rows[row - 1]) sheet.rows[row - 1].hidden = true;
    }
  }
  Object.entries(filter.criteria || {}).forEach(([colText, allowedValues]) => {
    const col = Number(colText);
    const criterion = normalizeAutoFilterCriterion(allowedValues);
    for (let row = range.top + 1; row <= range.bottom; row += 1) {
      if (!autoFilterCellMatchesCriterion(sheet, row, col, criterion) && sheet.rows[row - 1]) sheet.rows[row - 1].hidden = true;
    }
  });
  return true;
}



function applyStoredSortState(sheet, sortState) {
  const range = sortState?.range;
  const keys = Array.isArray(sortState?.keys) ? sortState.keys : [];
  if (!sheet || !range || !keys.length) return false;
  if (sortState.orientation === "columns") {
    return sortRangeColumnsByKeys(range, keys, Boolean(sortState.hasHeader), {
      sheet,
      orientation: "columns",
      caseSensitive: Boolean(sortState.caseSensitive),
      commitHistory: false,
      recordSortState: false,
      quiet: true,
    });
  }
  return sortRangeRowsByKeys(range, keys, Boolean(sortState.hasHeader), {
    sheet,
    caseSensitive: Boolean(sortState.caseSensitive),
    commitHistory: false,
    recordSortState: false,
    quiet: true,
  });
}



function autoFilterCellMatchesCriterion(sheet, row, col, criterion) {
  const normalized = normalizeAutoFilterCriterion(criterion);
  if (normalized.kind === "color") {
    const cell = sheet.cells[cellKey(row, col)];
    const css = cellEffectiveCss(sheet, cell);
    const actualColor =
      normalized.target === "font"
        ? cellHasFontColor(cell)
          ? normalizeSortColor(css.color)
          : ""
        : cellHasFill(cell)
          ? normalizeSortColor(css.backgroundColor)
          : "";
    return actualColor === normalizeSortColor(normalized.color);
  }
  return autoFilterValueMatchesCriterion(autoFilterCellValue(sheet, row, col), normalized);
}



function autoFilterVisibleRowStats(sheet, range = null) {
  const targetRange = range || sheet?.autoFilter?.range;
  if (!sheet || !targetRange) return { visible: 0, total: 0 };
  let visible = 0;
  const total = Math.max(0, targetRange.bottom - targetRange.top);
  for (let row = targetRange.top + 1; row <= targetRange.bottom; row += 1) {
    if (!sheet.rows[row - 1]?.hidden) visible += 1;
  }
  return { visible, total };
}



function autoFilterStatusMessage(message, sheet, range = null) {
  const stats = autoFilterVisibleRowStats(sheet, range);
  if (!stats.total) return `${message}。`;
  return `${message}（${stats.total.toLocaleString("ja-JP")} 行中 ${stats.visible.toLocaleString("ja-JP")} 行を表示）。`;
}



function openAutoFilterMenu(event) {
  event.preventDefault();
  event.stopPropagation();
  state.suppressNextCellClick = true;
  const button = event.currentTarget;
  const row = Number(button.dataset.filterRow);
  const col = Number(button.dataset.filterCol);
  const sheet = activeSheet();
  if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではオートフィルターを使用できません。")) return;
  if (!isAutoFilterHeaderCell(sheet, row, col)) return;
  activateAutoFilterForHeaderCell(sheet, row, col);
  closeAutoFilterMenu();
  const values = autoFilterUniqueValues(sheet, col);
  const criterion = normalizeAutoFilterCriterion(autoFilterStateForHeaderCell(sheet, row, col)?.criteria?.[col] || null);
  const activeValues = criterion.kind === "values" ? criterion.values : null;
  const customActive = criterion.kind === "custom";
  const colorActive = criterion.kind === "color";
  const allActive = criterion.kind === "all";
  const profile = autoFilterColumnProfile(sheet, col, values);
  const sortKey = autoFilterSortKeyForHeaderCell(sheet, row, col);
  const sortDirection = sortKey?.direction === "desc" ? "desc" : sortKey ? "asc" : "";
  const sortDescription = sortDirection
    ? sortKey.sortOn === "cellColor"
      ? "セルの色で並べ替えています。"
      : sortKey.sortOn === "fontColor"
        ? "フォントの色で並べ替えています。"
        : `${sortDirection === "desc" ? "降順" : "昇順"}で並べ替えています。`
    : "";
  const activeSet = new Set(activeValues || values.map((item) => item.value));
  const hasNarrowedValueFilter = criterion.kind === "values" && activeSet.size < values.length;
  const renderedValues = values.slice(0, AUTO_FILTER_VALUE_LIMIT);
  const hasMoreValues = values.length > renderedValues.length;
  const hasColumnCriterion = criterion.kind !== "all";
  const valueRows = renderedValues.map((item) => {
    const checked = allActive || activeSet.has(item.value) ? " checked" : "";
    const label = item.value === "" ? "(空白)" : item.value;
    return `<label class="auto-filter-value" data-filter-label="${escapeAttr(label.toLocaleLowerCase("ja-JP"))}"><input type="checkbox" data-auto-filter-value="${escapeAttr(item.value)}"${checked} />${escapeHtml(label)}</label>`;
  });
  const menu = document.createElement("div");
  menu.id = "autoFilterMenu";
  menu.className = "auto-filter-menu";
  menu.setAttribute("role", "menu");
  menu.dataset.filterCol = String(col);
  menu.dataset.hasValueFilter = hasNarrowedValueFilter ? "true" : "false";
  menu.dataset.hasColumnCriterion = hasColumnCriterion ? "true" : "false";
  menu.dataset.hasMoreValues = hasMoreValues ? "true" : "false";
  menu.dataset.totalValues = String(values.length);
  menu.dataset.renderedValues = String(renderedValues.length);
  menu.dataset.valueSelectionDirty = "false";
  menu.dataset.sortDirection = sortDirection;
  menu.innerHTML = `
    <button type="button" class="${sortDirection === "asc" ? "is-active" : ""}" data-auto-filter-action="sort-asc" aria-pressed="${sortDirection === "asc" ? "true" : "false"}">${sortDirection === "asc" ? "✓ " : ""}昇順</button>
    <button type="button" class="${sortDirection === "desc" ? "is-active" : ""}" data-auto-filter-action="sort-desc" aria-pressed="${sortDirection === "desc" ? "true" : "false"}">${sortDirection === "desc" ? "✓ " : ""}降順</button>
    <div class="auto-filter-separator"></div>
    ${autoFilterConditionMenuHtml(profile, customActive)}
    ${autoFilterColorMenuHtml(profile, colorActive)}
    <button type="button" data-auto-filter-action="clear"${hasColumnCriterion ? "" : " disabled"}>フィルターをクリア</button>
    <div class="auto-filter-current${hasColumnCriterion || sortDescription ? "" : " is-neutral"}">${escapeHtml([sortDescription, autoFilterCriterionDescription(criterion)].filter(Boolean).join(" "))}</div>
    ${hasMoreValues ? `<div class="auto-filter-limit-note">先頭 ${AUTO_FILTER_VALUE_LIMIT.toLocaleString("ja-JP")} 件だけ表示しています。検索または条件フィルターを使用してください。</div>` : ""}
    <div class="auto-filter-search">
      <input type="search" data-auto-filter-search placeholder="検索" aria-label="検索" />
      <label class="auto-filter-value auto-filter-add-current is-hidden"><input type="checkbox" data-auto-filter-add-current />現在の選択内容をフィルターに追加する</label>
      <label class="auto-filter-value"><input type="checkbox" data-auto-filter-select-all${allActive ? " checked" : ""} />(すべて選択)</label>
    </div>
    <div class="auto-filter-values">
      ${valueRows.join("")}
      <span class="auto-filter-empty${valueRows.length ? " is-hidden" : ""}" data-auto-filter-empty>表示する値がありません</span>
    </div>
    <div class="auto-filter-summary" data-auto-filter-menu-summary></div>
    <div class="auto-filter-actions">
      <button type="button" data-auto-filter-action="ok">OK</button>
      <button type="button" data-auto-filter-action="cancel">キャンセル</button>
    </div>`;
  document.body.appendChild(menu);
  const rect = button.getBoundingClientRect();
  menu.style.left = `${Math.min(window.innerWidth - menu.offsetWidth - 8, rect.left)}px`;
  menu.style.top = `${Math.min(window.innerHeight - menu.offsetHeight - 8, rect.bottom + 2)}px`;
  state.autoFilterMenu = { row, col };
  updateAutoFilterSelectAllState(menu);
  syncAutoFilterMenuState(menu);
  menu.querySelector("button, input")?.focus({ preventScroll: true });
  requestAnimationFrame(() => {
    if (!menu.contains(document.activeElement)) menu.querySelector("button, input")?.focus({ preventScroll: true });
  });
}



function openSelectedAutoFilterDropdown() {
  const sheet = activeSheet();
  const point = selectionToolbarPoint() || state.selected;
  if (!sheet || !point || point.sheetIndex !== state.activeSheetIndex) return false;
  let headerRow = point.row;
  const col = point.col;
  if (!isAutoFilterHeaderCell(sheet, headerRow, col)) {
    const context = autoFilterContextForSelection(sheet, activeSelectionRange());
    if (!context?.range || col < context.range.left || col > context.range.right) return false;
    headerRow = context.range.top;
  }
  if (!isAutoFilterHeaderCell(sheet, headerRow, col)) return false;

  const selector = `.cell-filter-button[data-filter-row="${headerRow}"][data-filter-col="${col}"]`;
  let button = document.querySelector(selector);
  if (!button) {
    scrollRangeIntoView({ sheetIndex: state.activeSheetIndex, top: headerRow, bottom: headerRow, left: col, right: col });
    renderSheet();
    button = document.querySelector(selector);
  }
  if (!button) return false;
  openAutoFilterMenu({
    preventDefault() {},
    stopPropagation() {},
    currentTarget: button,
  });
  return true;
}



function autoFilterCriterionDescription(criterion) {
  const normalized = normalizeAutoFilterCriterion(criterion);
  if (normalized.kind === "all") return "この列にフィルターは適用されていません。";
  if (normalized.kind === "values") {
    return `${(normalized.values || []).length.toLocaleString("ja-JP")} 項目を選択しています。`;
  }
  if (normalized.kind === "dateGroups") {
    const count = (normalized.dateGroups || []).length + (normalized.values || []).length + (normalized.includeBlank ? 1 : 0);
    return `${count.toLocaleString("ja-JP")} 件の日付項目を選択しています。`;
  }
  if (normalized.kind === "color") {
    const target = normalized.target === "font" ? "フォントの色" : "セルの色";
    return `${target}でフィルターしています。`;
  }
  if (normalized.top10) {
    const direction = normalized.top10.direction === "bottom" ? "下位" : "上位";
    const unit = normalized.top10.mode === "percent" ? "%" : "項目";
    return `トップ テン: ${direction} ${normalized.top10.count}${unit}。`;
  }
  if (normalized.averageOperator === "aboveAverage") return "平均より上でフィルターしています。";
  if (normalized.averageOperator === "belowAverage") return "平均より下でフィルターしています。";
  if (normalized.stdDevOperator === "aboveStdDev") return `平均より ${autoFilterStdDevCount(normalized.stdDevCount)} 標準偏差上でフィルターしています。`;
  if (normalized.stdDevOperator === "belowStdDev") return `平均より ${autoFilterStdDevCount(normalized.stdDevCount)} 標準偏差下でフィルターしています。`;
  if (normalized.dateGroup) return `${autoFilterDateGroupLabel(normalized.dateGroup)}でフィルターしています。`;
  if (normalized.datePeriod) return `${autoFilterDatePeriodLabel(normalized.datePeriod)}のすべての日付でフィルターしています。`;
  if (normalized.datePreset) return `${autoFilterDatePresetLabel(normalized.datePreset)}でフィルターしています。`;
  const first = autoFilterConditionDescription(normalized.operator1, normalized.value1);
  const hasSecond = String(normalized.value2 ?? "").trim() !== "" || autoFilterOperatorUsesBlankValue(normalized.operator2);
  if (!hasSecond) return `条件: ${first}`;
  const second = autoFilterConditionDescription(normalized.operator2, normalized.value2);
  return `条件: ${first} ${normalized.logic === "or" ? "または" : "かつ"} ${second}`;
}



function autoFilterConditionDescription(operator, value) {
  const text = String(value ?? "");
  switch (operator) {
    case "blank":
      return "空白セル";
    case "nonBlank":
      return "空白以外のセル";
    case "notEqual":
      return `「${text}」に等しくない`;
    case "contains":
      return `「${text}」を含む`;
    case "notContains":
      return `「${text}」を含まない`;
    case "beginsWith":
      return `「${text}」で始まる`;
    case "notBeginsWith":
      return `「${text}」で始まらない`;
    case "endsWith":
      return `「${text}」で終わる`;
    case "notEndsWith":
      return `「${text}」で終わらない`;
    case "greaterThan":
      return `「${text}」より大きい`;
    case "greaterThanOrEqual":
      return `「${text}」以上`;
    case "lessThan":
      return `「${text}」より小さい`;
    case "lessThanOrEqual":
      return `「${text}」以下`;
    case "equal":
    default:
      return `「${text}」に等しい`;
  }
}



function autoFilterDatePresetLabel(preset) {
  return {
    dateToday: "今日",
    dateYesterday: "昨日",
    dateTomorrow: "明日",
    dateThisWeek: "今週",
    dateLastWeek: "先週",
    dateNextWeek: "来週",
    dateThisMonth: "今月",
    dateLastMonth: "先月",
    dateNextMonth: "来月",
    dateThisQuarter: "今四半期",
    dateLastQuarter: "前四半期",
    dateNextQuarter: "次四半期",
    dateThisYear: "今年",
    dateLastYear: "昨年",
    dateNextYear: "来年",
    dateYearToDate: "年初来",
  }[preset] || "日付条件";
}



function autoFilterDatePeriodLabel(period) {
  const type = String(period || "");
  const month = type.match(/^M([1-9]|1[0-2])$/);
  if (month) return `${Number(month[1])}月`;
  const quarter = type.match(/^Q([1-4])$/);
  if (quarter) return `第${Number(quarter[1])}四半期`;
  return "期間内のすべての日付";
}



function autoFilterDateGroupLabel(group = {}) {
  const grouping = group.grouping || "year";
  const year = Number(group.year);
  const month = Number(group.month);
  const day = Number(group.day);
  const hour = Number(group.hour);
  const minute = Number(group.minute);
  const second = Number(group.second);
  if (!Number.isFinite(year)) return "日付";
  if (grouping === "year") return `${year}年`;
  if (!Number.isFinite(month)) return `${year}年`;
  if (grouping === "month") return `${year}年${month}月`;
  if (!Number.isFinite(day)) return `${year}年${month}月`;
  if (grouping === "day") return `${year}年${month}月${day}日`;
  if (!Number.isFinite(hour)) return `${year}年${month}月${day}日`;
  if (grouping === "hour") return `${year}年${month}月${day}日 ${String(hour).padStart(2, "0")}時`;
  if (!Number.isFinite(minute)) return `${year}年${month}月${day}日 ${String(hour).padStart(2, "0")}時`;
  if (grouping === "minute") return `${year}年${month}月${day}日 ${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
  if (!Number.isFinite(second)) return `${year}年${month}月${day}日 ${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
  return `${year}年${month}月${day}日 ${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}:${String(second).padStart(2, "0")}`;
}



function autoFilterColumnProfile(sheet, col, values = null) {
  const source = values || autoFilterUniqueValues(sheet, col);
  const nonBlank = source.map((item) => item.value).filter((value) => String(value).trim() !== "");
  const numericCount = nonBlank.filter((value) => Number.isFinite(Number(String(value).replace(/,/g, "")))).length;
  const dateCount = nonBlank.filter((value) => autoFilterDateSerial(value) != null).length;
  const numeric = numericCount > 0 && numericCount >= nonBlank.length - numericCount;
  const date = !numeric && dateCount > 0 && dateCount >= nonBlank.length - dateCount;
  return { numeric, date, nonBlankCount: nonBlank.length, ...autoFilterColumnColors(sheet, col) };
}



function autoFilterColumnColors(sheet, col) {
  const range = sheet?.autoFilter?.range;
  const fillColors = new Map();
  const fontColors = new Map();
  if (!range) return { fillColors: [], fontColors: [] };
  for (let row = range.top + 1; row <= range.bottom; row += 1) {
    const cell = sheet.cells[cellKey(row, col)];
    const css = cellEffectiveCss(sheet, cell);
    const fillColor = cellHasFill(cell) ? normalizeSortColor(css.backgroundColor) : "";
    const fontColor = cellHasFontColor(cell) ? normalizeSortColor(css.color) : "";
    if (fillColor) fillColors.set(fillColor, (fillColors.get(fillColor) || 0) + 1);
    if (fontColor) fontColors.set(fontColor, (fontColors.get(fontColor) || 0) + 1);
  }
  const toEntries = (map) =>
    Array.from(map.entries())
      .map(([color, count]) => ({ color, count }))
      .sort((left, right) => right.count - left.count || left.color.localeCompare(right.color));
  return { fillColors: toEntries(fillColors), fontColors: toEntries(fontColors) };
}

export {
  activateAutoFilterForHeaderCell,
  activateAutoFilterTable,
  activeWorkbookRefreshCount,
  applyAutoFilterStateToRows,
  applyQueryAuthenticationChanges,
  applyQueryDefinitionChanges,
  applyQueryPropertiesDialog,
  applyStoredSortState,
  autoFilterBlockRowHasData,
  autoFilterCellMatchesCriterion,
  autoFilterColumnColors,
  autoFilterColumnHasBodyData,
  autoFilterColumnProfile,
  autoFilterConditionDescription,
  autoFilterContextForSelection,
  autoFilterCriterionDescription,
  autoFilterDateGroupLabel,
  autoFilterDatePeriodLabel,
  autoFilterDatePresetLabel,
  autoFilterHeaderBlockBottom,
  autoFilterHeaderColumnHasSignal,
  autoFilterHeaderContinuousSpan,
  autoFilterHeaderProbeBottom,
  autoFilterHeaderRowBodyBottom,
  autoFilterHeaderRowRegionFromSelection,
  autoFilterHeaderSpanAt,
  autoFilterHeaderSpanHasSignal,
  autoFilterRangeFromSelection,
  autoFilterSortKeyForHeaderCell,
  autoFilterStateForHeaderCell,
  autoFilterStateHasCriteria,
  autoFilterStatusMessage,
  autoFilterTableForSelection,
  autoFilterTableRangeFromSelection,
  autoFilterVisibleRowStats,
  cancelWorkbookRefreshes,
  captureSortDialogLevels,
  cellHasDataForRegion,
  clearAutoFilterCriteria,
  clearAutoFilterCriteriaForSheetGroup,
  clearQueryPropertiesCredential,
  clearQueryRefreshTimers,
  clearSortStateFromMatchingTable,
  cloneColumnModelForMove,
  cloneRowModelForMove,
  comparableDateSerial,
  compareSortBlankValues,
  compareSortValues,
  connectionStringUrlForSourceType,
  connectionStringValue,
  currentRegionFromPoint,
  defaultSortDialogLevel,
  detectDelimitedTextDelimiter,
  firstUrlLikeText,
  handleQueryPropertiesDialogSync,
  handleSortDialogControlChange,
  inferHeaderRow,
  markWorkbookQueryFailed,
  markWorkbookQueryRefreshing,
  markWorkbookQuerySucceeded,
  mergedAutoFilterHeaderBlocks,
  mutateSortDialogLevels,
  normalizeCustomSortText,
  normalizeSortColor,
  normalizeSortDialogLevels,
  normalizeSortStateKeys,
  normalizedSortStateForColumns,
  normalizedSortStateForRows,
  oleDbConnectionInfo,
  openAutoFilterMenu,
  openRefreshStatusDialog,
  openSelectedAutoFilterDropdown,
  openSortDialog,
  parseDelimitedRows,
  parsePrefixedConnectionString,
  parseQueryConnectionString,
  queryAuthenticationChangesFromPropertiesDialog,
  queryDefinitionChangesFromPropertiesDialog,
  queryRefreshDurationLabel,
  queryRefreshDurationMs,
  queryRefreshPreflightError,
  queryRefreshTargetStillCurrent,
  querySourceForRefresh,
  querySourceItemNameFromCommandText,
  quickSortRangeFromSelection,
  rangeColumnHasData,
  rangeHasUserData,
  rangeRowHasData,
  readableRefreshError,
  reapplyAutoFilter,
  reapplyAutoFilterForSheetGroup,
  reapplyAutoFilterForSheetRefs,
  reapplyImportedAutoFiltersOnLoad,
  reapplyTableAutoFiltersOnLoad,
  reapplyVisibleTableAutoFilters,
  reapplyVisibleTableSorts,
  refreshScheduledQuery,
  refreshStatusDialogHtml,
  refreshStatusRowHtml,
  refreshWorkbookData,
  refreshWorkbookQueriesOnOpen,
  refreshWorkbookQueriesOnOpenAsync,
  refreshWorkbookQuery,
  refreshWorkbookQueryRemoteSource,
  refreshWorkbookQueryWithRemoteSource,
  removeAutoFilter,
  rerenderSortDialogLevels,
  resetRowsForAutoFilterRange,
  scheduleQueryRefreshTimers,
  setConnectionStringValue,
  setSortDialogActiveLevel,
  sheetUsedRange,
  shouldRemoveAutoFilterForSelection,
  shouldShowSortExpansionWarning,
  showDataOverwriteConfirmation,
  sortComparableColorValue,
  sortComparableValue,
  sortCustomListById,
  sortCustomListRank,
  sortDialogActiveIndex,
  sortDialogAxisLabel,
  sortDialogCellColor,
  sortDialogColorItems,
  sortDialogColorOptions,
  sortDialogCustomListOptions,
  sortDialogKeyOptions,
  sortDialogKeysFromLevels,
  sortDialogLevelGridHtml,
  sortDialogLevelRowHtml,
  sortDialogLevelsFromSortState,
  sortDialogMergeConflict,
  sortDialogOrderControlHtml,
  sortDialogOrderFromValue,
  sortDialogOrderLabel,
  sortDialogPreview,
  sortDialogSortOnLabel,
  sortDialogValidation,
  sortKeyForHeaderCell,
  sortRangeByActiveColumn,
  sortRangeColumnsByKeys,
  sortRangeRows,
  sortRangeRowsByKeys,
  sortSelectionByActiveColumn,
  sortSelectionByActiveColumnForSheetGroup,
  sortStateForDialog,
  sortStateHasKeys,
  syncAutoFilterStateToMatchingTable,
  syncQueryAuthenticationProperties,
  syncSortDialog,
  syncSortDialogOrientation,
  syncSortStateToMatchingTable,
  tableForAutoFilterHeaderCell,
  tableForAutoFilterRange,
  tableSortStateForAutoFilterContext,
  tableSortStateForFilterRange,
  textConnectionStringSource,
  toggleAutoFilterForSelection,
  toggleAutoFilterForSheetGroup,
  visibleTableAutoFilters,
  visibleTableSortStates,
};
