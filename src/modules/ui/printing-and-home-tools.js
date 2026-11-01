// Feature module; shared bindings are imported explicitly.
import {
  $formulaInput,
  $sheetHost,
  COLUMN_HEADER_HEIGHT,
  CUSTOM_SHAPE_TYPE,
  DATA_OPERATION_CELL_LIMIT,
  DRAW_COLOR_PRESETS,
  DRAW_INK_ERASERS,
  DRAW_INK_TOOLS,
  createIcons,
  icons,
  state,
} from "../../app-core.js";
import {
  closeWhatIfMenu,
} from "../data/analysis-and-scenarios.js";
import {
  closeQueryPaneContextMenu,
} from "../data/connections-and-queries.js";
import {
  closeAutoFilterMenu,
} from "../data/filter-and-text-tools.js";
import {
  closeFindDialog,
  findAllFromDialog,
  findNextFromDialog,
  matchingCellRanges,
  normalizeFindScope,
  replaceFromDialog,
  selectCellsByPredicate,
  selectFindResultFromDialog,
  selectRangesForHomeCommand,
  selectSingleRange,
} from "../data/query-editor.js";
import {
  closeDataCommandMenu,
} from "../data/query-transforms.js";
import {
  cellHasDataForRegion,
  clearAutoFilterCriteria,
  currentRegionFromPoint,
  openSortDialog,
  reapplyAutoFilter,
  sheetUsedRange,
  sortSelectionByActiveColumn,
  toggleAutoFilterForSelection,
} from "../data/sort-and-autofilter.js";
import {
  closeDataValidationMenu,
} from "../data/validation-and-consolidation.js";
import {
  accessKeyFromKeyboardEvent,
  activateDialogAccessKeyTarget,
  addShapeAtPixelBoxToSheet,
  cancelShapeDrawing,
  closeDataDialog,
  enhanceDialogAccessKeys,
  findDialogAccessKeyTarget,
} from "../drawing/charts-and-insertion.js";
import {
  clamp01,
  headerFooterSettingsFromSimpleText,
  normalizeShapeCustomPath,
  shapeDescriptorForType,
} from "../drawing/shape-rendering.js";
import {
  refreshShapeSource,
} from "../drawing/transforms-and-connectors.js";
import {
  applyAutoFillMerges,
  commitActiveCellEditForExternalPointer,
  setPointerCaptureSafely,
  shiftFormulaReferences,
  validateAutoFillMerges,
} from "../editing/autofill.js";
import {
  closeRibbonBorderMenu,
  closeRibbonColorPalette,
  closeRibbonInkEraserMenu,
  closeRibbonShapeMenu,
  closeRibbonUnderlineMenu,
  closeShapeFormatMenu,
  roundSvgNumber,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  isElementImeCompositionEvent,
} from "../editing/keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  affectedSheetIndexesForGroupedRanges,
  closeStatusBarContextMenu,
  cssEscape,
  expandRangeForMerges,
  getCellRawInput,
  getDisplayForCell,
  isSingleCellRange,
  materializedRangesForSheet,
  normalizeRange,
  rangeToLabel,
  rangesToLabel,
  refreshChartsForGroupedRanges,
  restoreGroupedSheetsAfterFailure,
  sheetPixelFromClientPosition,
  stopSelectionAutoScroll,
  updateSelectionUi,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  cellHasComment,
  cellHasConditionalFormat,
  cellHasConstant,
  cellHasFormula,
  cellHasNote,
  clearPictureCropState,
  closeNameBoxMenu,
  coordinateToPixels,
  imageBoxBaseStyle,
  imageBoxToPixels,
  imageObjectHtml,
  imageVisualBoundsToPixels,
  imageVisualBoxToPixels,
  inkStrokeSvg,
  objectVisualStyle,
  pixelBoxesIntersect,
  rawActiveSelectionRanges,
  roundCssNumber,
  selectedImageObjects,
  selectedInkStrokeIdSet,
  setSelectedImageIds,
  updateFormulaBarForSelection,
  worksheetObjectLayerZIndex,
} from "../editing/selection-and-cell-text.js";
import {
  captureCellsForHistory,
  captureCellsForHistoryRanges,
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneCellModel,
  cloneCellScalarValue,
  cloneJsonLikeValue,
  cloneMerges,
  internCellsInRange,
  internCellsInRanges,
  normalizeInkStrokes,
  pushHistory,
  stableObjectKey,
  storeCellModelForSheet,
} from "../editing/structure-and-history.js";
import {
  formulaReferenceTokensInText,
  openInsertFunctionDialog,
} from "../formula/dialogs-and-names.js";
import {
  ensureCellModel,
  forEachCellInRange,
  rangesIntersect,
  removeIntersectingMerges,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  closeCellContextMenu,
  positionPopupMenu,
} from "./annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  columnName,
  columnNumber,
  columnWidthToPixels,
  dateToExcelSerial,
  decodeRange,
  escapeAttr,
  escapeHtml,
  excelSerialToDate,
  setStatus,
  showMessageBox,
} from "./help-and-utilities.js";
import {
  applyAutoSumToSelection,
  closeFileSideMenu,
  ensureSheetSize,
  normalizeAutoSumFunctionName,
  parseCellKey,
  rowHeaderWidthForSheet,
  unquoteFormulaSheetName,
} from "./ribbon-and-links.js";
import {
  closeQuickAnalysisPopup,
  closeSheetNavigatorMenu,
  closeSheetTabContextMenu,
  ensureEditableRangeForProtection,
  ensureEditableRangesForProtection,
  ensureSheetProtectionPermission,
} from "./sheet-protection-and-translation.js";
import {
  setCalculationMode,
} from "./views-and-formula-tools.js";
import {
  buildFormulaEngine,
  ensureActiveSheetVisible,
  unquoteSheetName,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  activeSheet,
  displayValue,
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  decodeExcelRangeReference,
  firstDefinedNameReference,
  validationDefinedNameFromFormulaText,
  validationQualifiedRange,
  validationReferenceFromFormula,
} from "../workbook/tables-and-validation.js";
import {
  commitPageLayoutChange,
  normalizePageSetupPrintTitleColumns,
  normalizePageSetupPrintTitleRows,
  parsePageSetupMarginValue,
  printRenderScale,
  printRowHeight,
} from "../workbook/vba-and-printing.js";
import {
  closeViewCommandMenu,
  menuAccessKeyFromLabel,
  restoreWorksheetFocusAfterCommandPopup,
} from "../workbook/views-and-macros.js";


function printSheetObjectsHtml(sheet, range, layout, renderScale = 1) {
  if (sheet?.showObjects === false) return "";
  const origin = printRangePixelBounds(sheet, range);
  const scaleValue = printRenderScale(renderScale);
  const objects = [
    printSheetImagesHtml(sheet, range, layout, origin, scaleValue),
    printSheetInkHtml(sheet, range, layout, origin, scaleValue),
  ].filter(Boolean).join("");
  if (!objects) return "";
  return `<div class="workbook-print-object-layer" aria-hidden="true">${objects}</div>`;
}



function printRangePixelBounds(sheet, range) {
  const left = coordinateToPixels(sheet, "col", range.left - 1);
  const top = coordinateToPixels(sheet, "row", range.top - 1);
  const right = coordinateToPixels(sheet, "col", range.right);
  const bottom = coordinateToPixels(sheet, "row", range.bottom);
  return { left, top, right, bottom, width: Math.max(1, right - left), height: Math.max(1, bottom - top) };
}



function printHeadingOffsets(sheet, layout) {
  return {
    left: layout.printHeadings ? rowHeaderWidthForSheet(sheet) : 0,
    top: layout.printHeadings ? COLUMN_HEADER_HEIGHT : 0,
  };
}



function printSheetImagesHtml(sheet, range, layout, origin = printRangePixelBounds(sheet, range), renderScale = 1) {
  if (!sheet.images?.length) return "";
  const offsets = printHeadingOffsets(sheet, layout);
  const scaleValue = printRenderScale(renderScale);
  return sheet.images
    .map((image) => printImageHtml(sheet, image, origin, offsets, scaleValue))
    .filter(Boolean)
    .join("");
}



function printImageHtml(sheet, image, origin, offsets, renderScale = 1) {
  if (image?.hidden) return "";
  const box = imageBoxToPixels(sheet, image);
  if (!box) return "";
  const visualBounds = imageVisualBoundsToPixels(image, box);
  if (!pixelBoxesIntersect(visualBounds, origin)) return "";
  const scaleValue = printRenderScale(renderScale);
  const relativeBox = {
    left: (box.left - origin.left + offsets.left) * scaleValue,
    top: (box.top - origin.top + offsets.top) * scaleValue,
    width: box.width * scaleValue,
    height: box.height * scaleValue,
  };
  const style = `${imageBoxBaseStyle(image, relativeBox)}${objectVisualStyle(image)};position:absolute`;
  return imageObjectHtml(image, style, "", printObjectPositionAttrs(relativeBox), Number(sheet.displayScale || 1) * scaleValue);
}



function printObjectPositionAttrs(box) {
  return [
    `data-print-object="true"`,
    `data-print-object-left="${roundCssNumber(box.left)}"`,
    `data-print-object-top="${roundCssNumber(box.top)}"`,
    `data-print-object-width="${roundCssNumber(box.width)}"`,
    `data-print-object-height="${roundCssNumber(box.height)}"`,
  ].join(" ");
}



function syncPrintObjectLayerPositions(root = document) {
  if (!root?.querySelectorAll) return;
  root.querySelectorAll(".workbook-print-sheet-grid").forEach((grid) => {
    const rows = printRowAxisSegments(grid);
    const cols = printColumnAxisSegments(grid);
    if (!rows.length && !cols.length) return;
    grid.querySelectorAll("[data-print-object='true']").forEach((object) => {
      const left = Number(object.dataset.printObjectLeft);
      const top = Number(object.dataset.printObjectTop);
      const width = Number(object.dataset.printObjectWidth);
      const height = Number(object.dataset.printObjectHeight);
      if (![left, top, width, height].every(Number.isFinite)) return;
      const mappedLeft = cols.length ? mapPrintAxisCoordinate(cols, left) : left;
      const mappedRight = cols.length ? mapPrintAxisCoordinate(cols, left + width) : left + width;
      const mappedTop = rows.length ? mapPrintAxisCoordinate(rows, top) : top;
      const mappedBottom = rows.length ? mapPrintAxisCoordinate(rows, top + height) : top + height;
      object.style.left = `${roundCssNumber(mappedLeft)}px`;
      object.style.top = `${roundCssNumber(mappedTop)}px`;
      object.style.width = `${roundCssNumber(Math.max(1, mappedRight - mappedLeft))}px`;
      object.style.height = `${roundCssNumber(Math.max(1, mappedBottom - mappedTop))}px`;
    });
  });
}



function printRowAxisSegments(grid) {
  const rows = [...grid.querySelectorAll("tr[data-print-row-height]")];
  return printAxisSegments(rows.map((row) => ({
    expected: Number(row.dataset.printRowHeight) || row.offsetHeight || 1,
    actual: row.offsetHeight || Number(row.getBoundingClientRect?.().height) || Number(row.dataset.printRowHeight) || 1,
  })));
}



function printColumnAxisSegments(grid) {
  const cols = [...grid.querySelectorAll("col[data-print-col-width]")];
  const expectedWidths = cols.map((col) => Math.max(1, Number(col.dataset.printColWidth) || 1));
  if (!expectedWidths.length) return [];
  const expectedTotal = expectedWidths.reduce((sum, width) => sum + width, 0);
  const table = grid.querySelector(".workbook-print-table");
  const actualTotal = Math.max(1, Number(table?.offsetWidth) || expectedTotal);
  const ratio = actualTotal / Math.max(1, expectedTotal);
  return printAxisSegments(expectedWidths.map((width) => ({ expected: width, actual: width * ratio })));
}



function printAxisSegments(items) {
  let expectedStart = 0;
  let actualStart = 0;
  return items.map((item) => {
    const expectedSize = Math.max(1, Number(item.expected) || 1);
    const actualSize = Math.max(1, Number(item.actual) || expectedSize);
    const segment = {
      expectedStart,
      expectedEnd: expectedStart + expectedSize,
      actualStart,
      actualEnd: actualStart + actualSize,
      expectedSize,
      actualSize,
    };
    expectedStart += expectedSize;
    actualStart += actualSize;
    return segment;
  });
}



function mapPrintAxisCoordinate(segments, value) {
  const coordinate = Math.max(0, Number(value) || 0);
  if (!segments.length) return coordinate;
  const first = segments[0];
  if (coordinate <= first.expectedStart) return first.actualStart + (coordinate - first.expectedStart);
  for (const segment of segments) {
    if (coordinate <= segment.expectedEnd) {
      const ratio = Math.max(0, Math.min(1, (coordinate - segment.expectedStart) / segment.expectedSize));
      return segment.actualStart + ratio * segment.actualSize;
    }
  }
  const last = segments[segments.length - 1];
  return last.actualEnd + (coordinate - last.expectedEnd);
}



function printSheetInkHtml(sheet, range, layout, origin = printRangePixelBounds(sheet, range), renderScale = 1) {
  const strokes = normalizeInkStrokes(sheet?.inkStrokes);
  if (!strokes.length) return "";
  const offsets = printHeadingOffsets(sheet, layout);
  const scaleValue = printRenderScale(renderScale);
  const paths = strokes
    .filter((stroke) => pixelBoxesIntersect(inkStrokeBoundingBox(stroke), origin))
    .map((stroke) => inkStrokeSvg(stroke, false))
    .filter(Boolean)
    .join("");
  if (!paths) return "";
  const style = [
    `left:${roundCssNumber(offsets.left * scaleValue)}px`,
    `top:${roundCssNumber(offsets.top * scaleValue)}px`,
    `width:${roundCssNumber(origin.width * scaleValue)}px`,
    `height:${roundCssNumber(origin.height * scaleValue)}px`,
  ].join(";");
  return `<svg class="workbook-print-ink-layer" style="${escapeAttr(style)}" viewBox="${roundCssNumber(origin.left)} ${roundCssNumber(origin.top)} ${roundCssNumber(origin.width)} ${roundCssNumber(origin.height)}">${paths}</svg>`;
}



function handlePageSetupTabClick(event) {
  const tab = event.currentTarget;
  const dialog = tab.closest('[data-data-dialog="page-setup"]');
  if (!dialog) return;
  const id = tab.dataset.pageSetupTab;
  dialog.querySelectorAll("[data-page-setup-tab]").forEach((item) => item.classList.toggle("is-active", item === tab));
  dialog.querySelectorAll("[data-page-setup-panel]").forEach((panel) => panel.classList.toggle("is-active", panel.dataset.pageSetupPanel === id));
}



function applyPageSetupDialog(dialog) {
  const sheet = activeSheet();
  if (!sheet) return;
  const printAreaText = String(dialog.querySelector("[data-page-setup-field='printArea']")?.value || "").trim();
  const titleRowsText = String(dialog.querySelector("[data-page-setup-field='printTitleRows']")?.value || "").trim();
  const titleColumnsText = String(dialog.querySelector("[data-page-setup-field='printTitleColumns']")?.value || "").trim();
  const parsedPrintArea = printAreaText ? parsePageSetupRangeLabel(printAreaText, sheet) : null;
  const titleRows = titleRowsText ? parsePageSetupRangeLabel(titleRowsText, sheet) : null;
  const titleColumns = titleColumnsText ? parsePageSetupRangeLabel(titleColumnsText, sheet) : null;
  if (printAreaText && !parsedPrintArea) {
    showMessageBox("印刷範囲の参照が正しくありません。", { title: "Cht WebSheet" });
    return;
  }
  if (titleRowsText && !titleRows) {
    showMessageBox("タイトル行の参照が正しくありません。", { title: "Cht WebSheet" });
    return;
  }
  if (titleColumnsText && !titleColumns) {
    showMessageBox("タイトル列の参照が正しくありません。", { title: "Cht WebSheet" });
    return;
  }
  commitPageLayoutChange((layout) => {
    layout.orientation = dialog.querySelector("[data-page-setup-field='orientation']:checked")?.value || dialog.querySelector("[data-page-setup-field='orientation']")?.value || layout.orientation;
    layout.paperSize = dialog.querySelector("[data-page-setup-field='paperSize']")?.value || layout.paperSize;
    layout.scale = Number(dialog.querySelector("[data-page-setup-field='scale']")?.value) || layout.scale;
    layout.fitToWidth = dialog.querySelector("[data-page-setup-field='fitToWidth']")?.value || layout.fitToWidth;
    layout.fitToHeight = dialog.querySelector("[data-page-setup-field='fitToHeight']")?.value || layout.fitToHeight;
    layout.fitToPage = layout.fitToWidth !== "auto" || layout.fitToHeight !== "auto";
    layout.firstPageNumber = Number(dialog.querySelector("[data-page-setup-field='firstPageNumber']")?.value) || layout.firstPageNumber;
    layout.pageOrder = dialog.querySelector("[data-page-setup-field='pageOrder']")?.value || layout.pageOrder;
    layout.comments = dialog.querySelector("[data-page-setup-field='comments']")?.value || layout.comments;
    layout.cellErrors = dialog.querySelector("[data-page-setup-field='cellErrors']")?.value || layout.cellErrors;
    layout.printArea = parsedPrintArea ? { label: rangeToLabel(parsedPrintArea), range: parsedPrintArea } : null;
    layout.printTitles = {
      rows: normalizePageSetupPrintTitleRows(titleRows, sheet),
      cols: normalizePageSetupPrintTitleColumns(titleColumns, sheet),
    };
    layout.margins = { ...(layout.margins || {}) };
    dialog.querySelectorAll("[data-page-setup-margin]").forEach((input) => {
      const key = input.dataset.pageSetupMargin;
      layout.margins[key] = parsePageSetupMarginValue(input, key);
    });
    layout.printGridLines = Boolean(dialog.querySelector("[data-page-setup-check='printGridLines']")?.checked);
    layout.printHeadings = Boolean(dialog.querySelector("[data-page-setup-check='printHeadings']")?.checked);
    layout.centerHorizontal = Boolean(dialog.querySelector("[data-page-setup-check='centerHorizontal']")?.checked);
    layout.centerVertical = Boolean(dialog.querySelector("[data-page-setup-check='centerVertical']")?.checked);
    layout.blackAndWhite = Boolean(dialog.querySelector("[data-page-setup-check='blackAndWhite']")?.checked);
    layout.draftQuality = Boolean(dialog.querySelector("[data-page-setup-check='draftQuality']")?.checked);
    layout.headerFooter = {
      header: String(dialog.querySelector("[data-page-setup-field='headerText']")?.value || ""),
      footer: String(dialog.querySelector("[data-page-setup-field='footerText']")?.value || ""),
    };
    sheet.headerFooter = headerFooterSettingsFromSimpleText(layout.headerFooter.header, layout.headerFooter.footer, sheet.headerFooter);
  }, "ページ設定", "ページ設定を更新しました。");
  closeDataDialog();
}



function parsePageSetupRangeLabel(label, sheet = activeSheet()) {
  const parsed = pageSetupReferenceParts(label);
  if (!parsed?.text) return null;
  if (parsed.sheetName) {
    const expected = sheet?.name || activeSheet()?.name || "";
    if (expected && parsed.sheetName !== expected) return null;
  }
  const text = parsed.text.replace(/\$/g, "").trim().toUpperCase();
  const rowOnly = text.match(/^(\d+):(\d+)$/);
  if (rowOnly) {
    const top = Math.max(1, Math.min(Number(rowOnly[1]), Number(rowOnly[2])));
    const bottom = Math.max(top, Math.max(Number(rowOnly[1]), Number(rowOnly[2])));
    return { top, bottom, left: 1, right: sheet.colCount, sheetIndex: state.activeSheetIndex, fullRow: true };
  }
  const columnOnly = text.match(/^([A-Z]+):([A-Z]+)$/);
  if (columnOnly) {
    const left = Math.max(1, Math.min(columnNumber(columnOnly[1]), columnNumber(columnOnly[2])));
    const right = Math.max(left, Math.max(columnNumber(columnOnly[1]), columnNumber(columnOnly[2])));
    return { top: 1, bottom: sheet.rowCount, left, right, sheetIndex: state.activeSheetIndex, fullColumn: true };
  }
  const cell = (value) => {
    const match = String(value || "").trim().match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;
    return { col: Math.max(1, columnNumber(match[1])), row: Math.max(1, Number(match[2])) };
  };
  const [startText, endText = startText] = text.split(":");
  const start = cell(startText);
  const end = cell(endText);
  if (!start || !end) return null;
  return normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: start.row, col: start.col },
    { sheetIndex: state.activeSheetIndex, row: end.row, col: end.col },
  );
}



function pageSetupReferenceParts(label) {
  let text = String(label || "").trim();
  if (!text) return null;
  if (text.startsWith("=")) text = text.slice(1).trim();
  let sheetName = "";
  const bangIndex = pageSetupReferenceBangIndex(text);
  if (bangIndex >= 0) {
    const sheetText = text.slice(0, bangIndex).trim();
    text = text.slice(bangIndex + 1).trim();
    sheetName = unquoteFormulaSheetName(sheetText);
  }
  return { sheetName, text };
}



function pageSetupReferenceBangIndex(text) {
  let inQuote = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === "'") {
      if (inQuote && text[index + 1] === "'") {
        index += 1;
      } else {
        inQuote = !inQuote;
      }
    } else if (char === "!" && !inQuote) {
      return index;
    }
  }
  return -1;
}



function beginInkDrawingMode(toolKey) {
  cancelShapeDrawing({ silent: true });
  closeRibbonInkEraserMenu();
  if (toolKey === "lasso") {
    state.pendingInkTool = { type: "lasso", label: "なげなわ選択" };
    document.body.classList.add("is-drawing-ink");
    $sheetHost.addClass("is-drawing-ink");
    setStatus("なげなわ選択：オブジェクトを囲むようにドラッグしてください。");
    return;
  }
  const preset = DRAW_INK_TOOLS[toolKey] || DRAW_INK_TOOLS.pen;
  const selectedColor = state.currentInkColor || state.currentInkTool.color || preset.color;
  state.currentInkTool = { ...state.currentInkTool, ...preset, type: toolKey, color: selectedColor };
  state.pendingInkTool = { ...state.currentInkTool };
  document.body.classList.add("is-drawing-ink");
  $sheetHost.addClass("is-drawing-ink");
  syncDrawRibbonButtons();
  setStatus(`${preset.label}：シート上でドラッグして描画してください。`);
}



function beginInkEraserMode(eraserKey = state.currentInkEraser || "stroke") {
  cancelShapeDrawing({ silent: true });
  closeRibbonInkEraserMenu();
  const preset = DRAW_INK_ERASERS[eraserKey] || DRAW_INK_ERASERS.stroke;
  state.currentInkEraser = eraserKey;
  state.pendingInkTool = {
    type: "eraser",
    label: preset.label,
    eraserKey,
    eraserMode: preset.mode,
    eraserRadius: preset.radius,
  };
  document.body.classList.add("is-drawing-ink");
  $sheetHost.addClass("is-drawing-ink");
  syncDrawRibbonButtons();
  const detail = preset.mode === "stroke" ? "インクをクリックまたはドラッグして線全体を削除します。" : "ドラッグした部分だけを削除します。";
  setStatus(`${preset.label}：${detail}`);
}



function cancelInkDrawingMode() {
  state.pendingInkTool = null;
  if (state.inkDrawing?.preview) state.inkDrawing.preview.remove();
  state.inkDrawing = null;
  document.body.classList.remove("is-drawing-ink");
  $sheetHost.removeClass("is-drawing-ink");
  syncDrawRibbonButtons();
}



function setInkColor(colorKey) {
  const color = DRAW_COLOR_PRESETS[colorKey] || DRAW_COLOR_PRESETS.black;
  state.currentInkColor = color;
  state.currentInkTool = { ...state.currentInkTool, color };
  if (state.pendingInkTool && !["eraser", "lasso"].includes(state.pendingInkTool.type)) state.pendingInkTool.color = color;
  syncDrawRibbonButtons();
  setStatus("ペンの色を変更しました。");
}



function setInkWidth(width) {
  state.currentInkTool = { ...state.currentInkTool, width };
  if (state.pendingInkTool && !["eraser", "lasso"].includes(state.pendingInkTool.type)) state.pendingInkTool.width = width;
  setStatus(`ペンの太さを ${width}px に変更しました。`);
}



function syncDrawRibbonButtons() {
  const type = state.pendingInkTool?.type || "";
  ["draw-pen", "draw-pencil", "draw-highlighter", "draw-eraser", "draw-lasso"].forEach((command) => {
    document.querySelector(`[data-ribbon-command="${command}"]`)?.classList.toggle("is-active", command === `draw-${type}`);
  });
  const eraserPreset = DRAW_INK_ERASERS[state.currentInkEraser] || DRAW_INK_ERASERS.stroke;
  const eraserButton = document.querySelector('[data-ribbon-command="draw-eraser"]');
  if (eraserButton) {
    eraserButton.title = eraserPreset.label;
    const label = eraserButton.querySelector("span");
    if (label) label.textContent = eraserPreset.shortLabel || "消しゴム";
  }
  Object.entries(DRAW_COLOR_PRESETS).forEach(([key, color]) => {
    document.querySelector(`[data-ribbon-command="draw-color-${key}"]`)?.classList.toggle("is-active", color.toLowerCase() === String(state.currentInkTool.color || "").toLowerCase());
  });
}



function startInkDrawing(event) {
  const tool = state.pendingInkTool;
  if (!tool || event.button !== 0) return false;
  if (tool.type === "eraser") {
    const sheet = activeSheet();
    if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではインクを削除できません。")) return false;
    closeRibbonShapeMenu();
    closeRibbonInkEraserMenu();
    stopSelectionAutoScroll();
    state.inkDrawing = {
      tool: { ...tool },
      pointerId: event.pointerId,
      sheetIndex: state.activeSheetIndex,
      eraserBefore: captureSheetForHistory(sheet),
      eraserChanged: false,
      selectionBefore: activeSelectionRange(),
      selectionBeforeRanges: activeSelectionRanges(),
    };
    eraseInkAtClientPoint(event.clientX, event.clientY, { deferHistory: true });
    try {
      setPointerCaptureSafely(event.currentTarget, event.pointerId);
    } catch {}
    event.preventDefault();
    event.stopPropagation();
    return true;
  }
  const point = sheetPixelFromClientPosition(event.clientX, event.clientY);
  if (!point) return false;
  closeRibbonShapeMenu();
  stopSelectionAutoScroll();
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectedImage = null;
  state.selectedImages = [];
  state.isSelecting = false;
  state.inkDrawing = {
    tool: { ...tool },
    points: [point],
    pointerId: event.pointerId,
    preview: createInkDrawingPreview(tool),
  };
  $sheetHost.find(".sheet-grid")[0]?.append(state.inkDrawing.preview);
  updateInkDrawingPreview();
  try {
    event.currentTarget.setPointerCapture?.(event.pointerId);
  } catch {}
  event.preventDefault();
  event.stopPropagation();
  return true;
}



function updateInkDrawing(event) {
  const drawing = state.inkDrawing;
  if (!drawing) return;
  if (drawing.pointerId && event.pointerId && drawing.pointerId !== event.pointerId) return;
  if (drawing.tool.type === "eraser") {
    eraseInkAtClientPoint(event.clientX, event.clientY, { deferHistory: true });
    event.preventDefault?.();
    return;
  }
  const point = sheetPixelFromClientPosition(event.clientX, event.clientY, { clampToGrid: true });
  if (!point) return;
  const last = drawing.points.at(-1);
  if (!last || Math.hypot(point.x - last.x, point.y - last.y) >= 2) {
    drawing.points.push(point);
    updateInkDrawingPreview();
  }
  event.preventDefault?.();
}



function finishInkDrawing(event) {
  const drawing = state.inkDrawing;
  if (!drawing) return;
  if (drawing.pointerId && event.pointerId && drawing.pointerId !== event.pointerId) return;
  const tool = drawing.tool;
  if (tool.type === "eraser") {
    state.inkDrawing = null;
    finishInkEraser(drawing);
    event.preventDefault?.();
    return;
  }
  const points = drawing.points.slice();
  drawing.preview?.remove();
  state.inkDrawing = null;
  if (tool.type === "lasso") selectInkObjectsInsideLasso(points);
  else insertInkPathFromPoints(points, tool);
  event.preventDefault?.();
}



function finishInkEraser(drawing) {
  const sheetIndex = Number.isInteger(drawing?.sheetIndex) ? drawing.sheetIndex : state.activeSheetIndex;
  const sheet = state.model?.sheets?.[sheetIndex] || activeSheet();
  if (!drawing?.eraserChanged || !sheet) {
    setStatus("削除するインクがありません。");
    return;
  }
  pushHistory({
    sheetIndex,
    sheetBefore: drawing.eraserBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: drawing.selectionBefore || null,
    selectionAfter: null,
    selectionBeforeRanges: drawing.selectionBeforeRanges || [],
    selectionAfterRanges: [],
    label: "インクの削除",
  });
  setStatus("インクを削除しました。");
}



function createInkDrawingPreview(tool) {
  const preview = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  preview.classList.add("sheet-ink-draft");
  preview.setAttribute("aria-hidden", "true");
  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("fill", "none");
  path.setAttribute("stroke", tool.color || "#111827");
  path.setAttribute("stroke-width", String(tool.width || 2));
  path.setAttribute("stroke-linecap", "round");
  path.setAttribute("stroke-linejoin", "round");
  path.setAttribute("opacity", String(tool.opacity ?? 1));
  preview.append(path);
  return preview;
}



function updateInkDrawingPreview() {
  const drawing = state.inkDrawing;
  if (!drawing?.preview) return;
  const box = pointListBoundingBox(drawing.points, Math.max(4, drawing.tool.width || 2));
  drawing.preview.setAttribute("viewBox", `${box.left} ${box.top} ${box.width} ${box.height}`);
  Object.assign(drawing.preview.style, {
    left: `${box.left}px`,
    top: `${box.top}px`,
    width: `${box.width}px`,
    height: `${box.height}px`,
  });
  drawing.preview.querySelector("path")?.setAttribute("d", svgPathDFromSheetPoints(drawing.points));
}



function insertInkPathFromPoints(points, tool) {
  if (!points || points.length < 2) return;
  const sheet = activeSheet();
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではインクを描画できません。")) return;
  const stroke = createInkStrokeFromPoints(points, tool);
  if (!stroke) return;
  const before = captureSheetForHistory(sheet);
  sheet.inkStrokes = normalizeInkStrokes(sheet.inkStrokes);
  sheet.inkStrokes.push(stroke);
  state.selectedInkStrokeIds = [];
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore: before,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${stroke.name || "インク"}の描画`,
  });
  renderSheet();
  updateSelectionUi();
  setStatus(`${stroke.name || "インク"}を描画しました。`);
}



function createInkStrokeFromPoints(points, tool = {}) {
  const sampled = simplifyInkPoints(points, 800).map((point) => ({
    x: roundCssNumber(point.x),
    y: roundCssNumber(point.y),
  }));
  if (sampled.length < 2) return null;
  return normalizeInkStrokes([{
    id: uniqueWorksheetObjectId("ink"),
    name: tool.label || "インク",
    tool: tool.type || "pen",
    color: tool.color || "#111827",
    width: tool.width || 2,
    opacity: tool.opacity ?? 1,
    points: sampled,
  }])[0] || null;
}



function inkStrokeColor(tool = {}) {
  const color = tool.color || "#111827";
  const opacity = Number(tool.opacity);
  if (!Number.isFinite(opacity) || opacity >= 0.999 || !/^#[0-9a-f]{6}$/i.test(color)) return color;
  const r = Number.parseInt(color.slice(1, 3), 16);
  const g = Number.parseInt(color.slice(3, 5), 16);
  const b = Number.parseInt(color.slice(5, 7), 16);
  return `rgba(${r},${g},${b},${Math.max(0.05, Math.min(1, opacity))})`;
}



function pointListBoundingBox(points, padding = 4) {
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const left = Math.max(rowHeaderWidthForSheet(activeSheet()), Math.min(...xs) - padding);
  const top = Math.max(COLUMN_HEADER_HEIGHT, Math.min(...ys) - padding);
  const right = Math.max(...xs) + padding;
  const bottom = Math.max(...ys) + padding;
  return {
    left,
    top,
    width: Math.max(12, right - left),
    height: Math.max(12, bottom - top),
  };
}



function inkStrokeBoundingBox(stroke, extraPadding = 0) {
  const points = Array.isArray(stroke?.points) ? stroke.points : [];
  if (!points.length) return { left: 0, top: 0, width: 0, height: 0 };
  const padding = Math.max(2, Number(stroke.width) || 2) / 2 + Math.max(0, Number(extraPadding) || 0);
  const xs = points.map((point) => Number(point.x)).filter(Number.isFinite);
  const ys = points.map((point) => Number(point.y)).filter(Number.isFinite);
  if (!xs.length || !ys.length) return { left: 0, top: 0, width: 0, height: 0 };
  const left = Math.min(...xs) - padding;
  const top = Math.min(...ys) - padding;
  const right = Math.max(...xs) + padding;
  const bottom = Math.max(...ys) + padding;
  return {
    left,
    top,
    width: Math.max(1, right - left),
    height: Math.max(1, bottom - top),
  };
}



function customPathFromSheetPoints(points, box) {
  const sampled = simplifyInkPoints(points, 120);
  const commands = sampled.map((point, index) => ({
    cmd: index === 0 ? "M" : "L",
    points: [{
      x: clamp01((point.x - box.left) / Math.max(1, box.width)),
      y: clamp01((point.y - box.top) / Math.max(1, box.height)),
    }],
  }));
  return normalizeShapeCustomPath({ commands });
}



function simplifyInkPoints(points, maxPoints) {
  if (points.length <= maxPoints) return points;
  const step = Math.ceil(points.length / maxPoints);
  const sampled = points.filter((_point, index) => index % step === 0);
  if (sampled.at(-1) !== points.at(-1)) sampled.push(points.at(-1));
  return sampled;
}



function svgPathDFromSheetPoints(points) {
  if (!points.length) return "";
  return points.map((point, index) => `${index === 0 ? "M" : "L"}${roundSvgNumber(point.x)} ${roundSvgNumber(point.y)}`).join(" ");
}



function eraseInkAtClientPoint(clientX, clientY, options = {}) {
  const point = sheetPixelFromClientPosition(clientX, clientY, { clampToGrid: true });
  if (!point) return;
  const sheet = activeSheet();
  const eraserTool = state.inkDrawing?.tool?.type === "eraser" ? state.inkDrawing.tool : state.pendingInkTool;
  const preset = DRAW_INK_ERASERS[eraserTool?.eraserKey] || DRAW_INK_ERASERS[state.currentInkEraser] || DRAW_INK_ERASERS.stroke;
  const mode = eraserTool?.eraserMode || preset.mode || "stroke";
  const radius = Math.max(2, Number(eraserTool?.eraserRadius ?? preset.radius) || preset.radius || 8);
  const eraseResult = mode === "point"
    ? eraseInkPointAtSheetPoint(sheet, point, radius)
    : eraseInkStrokeAtSheetPoint(sheet, point, radius);
  if (!eraseResult.changed) return;
  const before = state.inkDrawing?.tool?.type === "eraser" || options.deferHistory
    ? null
    : eraseResult.before;
  setSelectedImageIds([], null);
  if (state.inkDrawing?.tool?.type === "eraser" || options.deferHistory) {
    if (state.inkDrawing) state.inkDrawing.eraserChanged = true;
  } else {
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetBefore: before,
      sheetAfter: captureSheetForHistory(sheet),
      selectionBefore: activeSelectionRange(),
      selectionAfter: null,
      selectionBeforeRanges: activeSelectionRanges(),
      selectionAfterRanges: [],
      label: "インクの削除",
    });
  }
  renderSheet();
  setStatus("インクを削除しました。");
}



function eraseInkStrokeAtSheetPoint(sheet, point, radius = 8) {
  const targetStroke = hitTestInkStroke(sheet, point, Math.max(4, radius));
  const targetLegacyImage = targetStroke ? null : [...(sheet.images || [])].reverse().find((image) => image?.shape?.ink && pixelPointInsideImage(sheet, image, point, radius));
  if (!targetStroke && !targetLegacyImage) return { changed: false, before: null };
  const before = captureSheetForHistory(sheet);
  if (targetStroke) {
    sheet.inkStrokes = normalizeInkStrokes(sheet.inkStrokes).filter((stroke) => stroke.id !== targetStroke.id);
    state.selectedInkStrokeIds = (state.selectedInkStrokeIds || []).filter((id) => id !== targetStroke.id);
  } else {
    sheet.images = sheet.images.filter((image) => image.id !== targetLegacyImage.id);
  }
  return { changed: true, before };
}



function eraseInkPointAtSheetPoint(sheet, point, radius = 10) {
  const strokes = normalizeInkStrokes(sheet.inkStrokes);
  const nextStrokes = [];
  let changed = false;
  const changedIds = new Set();
  strokes.forEach((stroke) => {
    const effectiveRadius = Math.max(2, radius + Math.max(0.5, Number(stroke.width) || 2) / 2);
    const pointBox = { left: point.x - effectiveRadius, top: point.y - effectiveRadius, width: effectiveRadius * 2, height: effectiveRadius * 2 };
    if (!pixelBoxesIntersect(inkStrokeBoundingBox(stroke, effectiveRadius), pointBox)) {
      nextStrokes.push(stroke);
      return;
    }
    const result = eraseInkStrokePartAtPoint(stroke, point, effectiveRadius);
    if (!result.changed) {
      nextStrokes.push(stroke);
      return;
    }
    changed = true;
    changedIds.add(stroke.id);
    nextStrokes.push(...result.strokes);
  });
  const targetLegacyImage = changed ? null : [...(sheet.images || [])].reverse().find((image) => image?.shape?.ink && pixelPointInsideImage(sheet, image, point, radius));
  if (!changed && !targetLegacyImage) return { changed: false, before: null };
  const before = captureSheetForHistory(sheet);
  if (changed) {
    sheet.inkStrokes = normalizeInkStrokes(nextStrokes);
    const remainingIds = new Set(sheet.inkStrokes.map((stroke) => stroke.id));
    state.selectedInkStrokeIds = (state.selectedInkStrokeIds || []).filter((id) => remainingIds.has(id) && !changedIds.has(id));
  }
  if (targetLegacyImage) {
    sheet.images = sheet.images.filter((image) => image.id !== targetLegacyImage.id);
  }
  return { changed: true, before };
}



function eraseInkStrokePartAtPoint(stroke, center, radius) {
  const points = (stroke?.points || [])
    .map((point) => ({ x: Number(point.x), y: Number(point.y) }))
    .filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
  if (points.length < 2) return { changed: false, strokes: [stroke] };
  const runs = [];
  let currentRun = [];
  let touched = false;
  const finishRun = () => {
    if (currentRun.length >= 2) runs.push(currentRun);
    currentRun = [];
  };
  const pushRunPoint = (point) => {
    const rounded = { x: roundCssNumber(point.x), y: roundCssNumber(point.y) };
    const last = currentRun.at(-1);
    if (!last || Math.hypot(last.x - rounded.x, last.y - rounded.y) > 0.35) currentRun.push(rounded);
  };
  for (let index = 1; index < points.length; index += 1) {
    const start = points[index - 1];
    const end = points[index];
    const startInside = Math.hypot(start.x - center.x, start.y - center.y) <= radius;
    const endInside = Math.hypot(end.x - center.x, end.y - center.y) <= radius;
    const segmentTouched = startInside || endInside || distancePointToSegment(center, start, end) <= radius;
    if (!segmentTouched) {
      if (!currentRun.length) pushRunPoint(start);
      pushRunPoint(end);
      continue;
    }
    touched = true;
    const intersections = segmentCircleIntersections(start, end, center, radius);
    if (!startInside && !currentRun.length) pushRunPoint(start);
    if (!startInside && endInside) {
      const entry = intersections[0];
      if (Number.isFinite(entry)) pushRunPoint(interpolateInkPoint(start, end, entry));
      finishRun();
      continue;
    }
    if (startInside && endInside) {
      finishRun();
      continue;
    }
    if (startInside && !endInside) {
      finishRun();
      const exit = intersections.at(-1);
      if (Number.isFinite(exit)) pushRunPoint(interpolateInkPoint(start, end, exit));
      pushRunPoint(end);
      continue;
    }
    if (intersections.length >= 2) {
      pushRunPoint(interpolateInkPoint(start, end, intersections[0]));
      finishRun();
      pushRunPoint(interpolateInkPoint(start, end, intersections[1]));
      pushRunPoint(end);
      continue;
    }
    pushRunPoint(end);
  }
  finishRun();
  if (!touched) return { changed: false, strokes: [stroke] };
  const strokes = runs.map((points, index) => normalizeInkStrokes([{
    ...stroke,
    id: index === 0 ? stroke.id : uniqueWorksheetObjectId("ink"),
    points,
  }])[0]).filter(Boolean);
  return { changed: true, strokes };
}



function segmentCircleIntersections(start, end, center, radius) {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const a = dx * dx + dy * dy;
  if (a <= 0) return [];
  const fx = start.x - center.x;
  const fy = start.y - center.y;
  const b = 2 * (fx * dx + fy * dy);
  const c = fx * fx + fy * fy - radius * radius;
  const discriminant = b * b - 4 * a * c;
  if (discriminant < 0) return [];
  const root = Math.sqrt(Math.max(0, discriminant));
  const values = [(-b - root) / (2 * a), (-b + root) / (2 * a)]
    .filter((value) => value > 0 && value < 1)
    .sort((left, right) => left - right);
  return values.filter((value, index) => index === 0 || Math.abs(value - values[index - 1]) > 0.001);
}



function interpolateInkPoint(start, end, t) {
  return {
    x: start.x + (end.x - start.x) * t,
    y: start.y + (end.y - start.y) * t,
  };
}



function hitTestInkStroke(sheet, point, tolerance = 6) {
  const strokes = normalizeInkStrokes(sheet?.inkStrokes);
  for (let index = strokes.length - 1; index >= 0; index -= 1) {
    const stroke = strokes[index];
    const box = inkStrokeBoundingBox(stroke, tolerance);
    if (!pixelBoxesIntersect(box, { left: point.x, top: point.y, width: 1, height: 1 })) continue;
    const threshold = Math.max(tolerance, (Number(stroke.width) || 2) / 2 + tolerance);
    if (inkStrokeDistanceToPoint(stroke, point) <= threshold) return stroke;
  }
  return null;
}



function inkStrokeDistanceToPoint(stroke, point) {
  const points = stroke?.points || [];
  if (!points.length) return Number.POSITIVE_INFINITY;
  let best = Number.POSITIVE_INFINITY;
  for (let index = 1; index < points.length; index += 1) {
    best = Math.min(best, distancePointToSegment(point, points[index - 1], points[index]));
  }
  return best;
}



function distancePointToSegment(point, start, end) {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const lengthSq = dx * dx + dy * dy;
  if (!lengthSq) return Math.hypot(point.x - start.x, point.y - start.y);
  const t = Math.max(0, Math.min(1, ((point.x - start.x) * dx + (point.y - start.y) * dy) / lengthSq));
  const x = start.x + dx * t;
  const y = start.y + dy * t;
  return Math.hypot(point.x - x, point.y - y);
}



function pixelPointInsideImage(sheet, image, point, padding = 0) {
  const box = imageVisualBoxToPixels(image, imageBoxToPixels(sheet, image));
  return Boolean(box && point.x >= box.left - padding && point.x <= box.left + box.width + padding && point.y >= box.top - padding && point.y <= box.top + box.height + padding);
}



function selectInkObjectsInsideLasso(points) {
  if (!points || points.length < 2) return;
  const sheet = activeSheet();
  const box = pointListBoundingBox(points, 0);
  const strokeIds = normalizeInkStrokes(sheet.inkStrokes)
    .filter((stroke) => {
      const strokeBox = inkStrokeBoundingBox(stroke);
      const center = { x: strokeBox.left + strokeBox.width / 2, y: strokeBox.top + strokeBox.height / 2 };
      return center.x >= box.left && center.x <= box.left + box.width && center.y >= box.top && center.y <= box.top + box.height;
    })
    .map((stroke) => stroke.id);
  const legacyIds = (sheet.images || [])
    .filter((image) => image?.shape?.ink)
    .filter((image) => {
      const imageBox = imageVisualBoxToPixels(image, imageBoxToPixels(sheet, image));
      if (!imageBox) return false;
      const center = { x: imageBox.left + imageBox.width / 2, y: imageBox.top + imageBox.height / 2 };
      return center.x >= box.left && center.x <= box.left + box.width && center.y >= box.top && center.y <= box.top + box.height;
    })
    .map((image) => image.id);
  if (!strokeIds.length && !legacyIds.length) {
    setStatus("なげなわ範囲にインクがありません。");
    return;
  }
  setSelectedImageIds(legacyIds, legacyIds[0]);
  state.selectedInkStrokeIds = strokeIds;
  renderSheet();
  updateSelectionUi();
  setStatus(`${strokeIds.length + legacyIds.length} 個のインクを選択しました。`);
}



function convertSelectedInkToFreeformShape() {
  const sheet = activeSheet();
  const selectedStrokeIds = selectedInkStrokeIdSet();
  const strokes = normalizeInkStrokes(sheet.inkStrokes).filter((stroke) => selectedStrokeIds.has(stroke.id));
  const legacyImages = selectedImageObjects().filter((image) => image?.shape?.ink);
  if (!strokes.length && !legacyImages.length) {
    setStatus("変換するインクを選択してください。");
    return;
  }
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではインクを図形に変換できません。")) return;
  const before = captureSheetForHistory(sheet);
  const insertedIds = [];
  strokes.forEach((stroke) => {
    const box = inkStrokeBoundingBox(stroke, Math.max(2, stroke.width || 2));
    const customPath = customPathFromSheetPoints(stroke.points, box);
    if (!customPath) return;
    const descriptor = shapeDescriptorForType(CUSTOM_SHAPE_TYPE, {
      width: box.width,
      height: box.height,
      customPath,
      fillColor: "none",
      outlineColor: inkStrokeColor(stroke),
      outlineWidth: stroke.width || 2,
      lineCap: "round",
      lineJoin: "round",
    });
    descriptor.name = stroke.name?.replace(/^インク/, "自由図形") || "自由図形";
    descriptor.shape = {
      ...descriptor.shape,
      ink: false,
      label: "自由図形",
      fillColor: "none",
    };
    const image = addShapeAtPixelBoxToSheet(sheet, CUSTOM_SHAPE_TYPE, box, descriptor);
    if (!image) return;
    image.name = descriptor.name;
    image.src = descriptor.src;
    image.shape = descriptor.shape;
    refreshShapeSource(image);
    insertedIds.push(image.id);
  });
  sheet.inkStrokes = normalizeInkStrokes(sheet.inkStrokes).filter((stroke) => !selectedStrokeIds.has(stroke.id));
  legacyImages.forEach((image) => {
    image.name = image.name?.replace(/^インク/, "自由図形") || "自由図形";
    image.shape = { ...(image.shape || {}), ink: false, label: "自由図形", fillColor: "none" };
    refreshShapeSource(image);
    insertedIds.push(image.id);
  });
  state.selectedInkStrokeIds = [];
  setSelectedImageIds(insertedIds, insertedIds[0]);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore: before,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionAfter: null,
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: [],
    label: "インクを図形に変換",
  });
  renderSheet();
  updateSelectionUi();
  setStatus(`${strokes.length + legacyImages.length} 個のインクを図形に変換しました。`);
}



function clearInkObjects() {
  const sheet = activeSheet();
  const strokes = normalizeInkStrokes(sheet.inkStrokes);
  const targets = (sheet.images || []).filter((image) => image?.shape?.ink);
  if (!strokes.length && !targets.length) {
    setStatus("削除するインクがありません。");
    return;
  }
  const before = captureSheetForHistory(sheet);
  const ids = new Set(targets.map((image) => image.id));
  sheet.inkStrokes = [];
  sheet.images = sheet.images.filter((image) => !ids.has(image.id));
  state.selectedInkStrokeIds = [];
  setSelectedImageIds([], null);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore: before,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: "インクの削除",
  });
  renderSheet();
  setStatus(`${strokes.length + targets.length} 個のインクを削除しました。`);
}



function autoSumOperationsForRanges(sheet, sheetId, ranges, functionName = "SUM", sheetIndex = state.activeSheetIndex) {
  const operations = [];
  const formulaName = normalizeAutoSumFunctionName(functionName);
  (ranges.length ? ranges : [activeSelectionRange()]).filter(Boolean).forEach((range) => {
    if (isSingleCellRange(range)) {
      const formula = autoSumFormulaForCell(sheet, sheetId, range.top, range.left, formulaName, sheetIndex);
      if (formula) operations.push({ target: normalizeRange({ sheetIndex, row: range.top, col: range.left }, { sheetIndex, row: range.top, col: range.left }), formula });
      return;
    }

    const width = range.right - range.left + 1;
    const height = range.bottom - range.top + 1;
    if (height >= width) {
      const bottomBlank = rangeBottomRowIsBlank(sheet, range);
      const targetRow = bottomBlank ? range.bottom : range.bottom + 1;
      const sourceTop = range.top;
      const sourceBottom = bottomBlank ? range.bottom - 1 : range.bottom;
      for (let col = range.left; col <= range.right; col += 1) {
        if (sourceBottom < sourceTop) continue;
        const source = normalizeRange({ sheetIndex, row: sourceTop, col }, { sheetIndex, row: sourceBottom, col });
        if (!rangeHasSummableCell(sheet, sheetId, source)) continue;
        operations.push({
          target: normalizeRange({ sheetIndex, row: targetRow, col }, { sheetIndex, row: targetRow, col }),
          formula: `=${formulaName}(${rangeToLabel(source)})`,
        });
      }
      return;
    }

    const rightBlank = rangeRightColumnIsBlank(sheet, range);
    const targetCol = rightBlank ? range.right : range.right + 1;
    const sourceLeft = range.left;
    const sourceRight = rightBlank ? range.right - 1 : range.right;
    for (let row = range.top; row <= range.bottom; row += 1) {
      if (sourceRight < sourceLeft) continue;
      const source = normalizeRange({ sheetIndex, row, col: sourceLeft }, { sheetIndex, row, col: sourceRight });
      if (!rangeHasSummableCell(sheet, sheetId, source)) continue;
      operations.push({
        target: normalizeRange({ sheetIndex, row, col: targetCol }, { sheetIndex, row, col: targetCol }),
        formula: `=${formulaName}(${rangeToLabel(source)})`,
      });
    }
  });
  return operations;
}



function autoSumFormulaForCell(sheet, sheetId, row, col, functionName = "SUM", sheetIndex = state.activeSheetIndex) {
  const formulaName = normalizeAutoSumFunctionName(functionName);
  const vertical = contiguousAutoSumSource(sheet, sheetId, row - 1, col, -1, 0, sheetIndex);
  if (vertical) return `=${formulaName}(${rangeToLabel(vertical)})`;
  const horizontal = contiguousAutoSumSource(sheet, sheetId, row, col - 1, 0, -1, sheetIndex);
  if (horizontal) return `=${formulaName}(${rangeToLabel(horizontal)})`;
  return "";
}



function contiguousAutoSumSource(sheet, sheetId, startRow, startCol, rowStep, colStep, sheetIndex = state.activeSheetIndex) {
  let row = startRow;
  let col = startCol;
  let first = null;
  let last = null;
  while (row >= 1 && col >= 1 && autoSumCellLooksSummable(sheet, sheetId, row, col)) {
    const point = { sheetIndex, row, col };
    if (!first) first = point;
    last = point;
    row += rowStep;
    col += colStep;
  }
  if (!first || !last) return null;
  return normalizeRange(first, last);
}



function autoSumCellLooksSummable(sheet, sheetId, row, col) {
  const cell = sheet.cells[cellKey(row, col)];
  if (!cell) return false;
  if (cell.formula || (typeof cell.raw === "string" && cell.raw.startsWith("="))) return true;
  const value = cell.raw ?? cell.cached ?? getDisplayForCell(sheet, sheetId, row, col);
  if (value == null || String(value).trim() === "") return false;
  if (typeof value === "number") return Number.isFinite(value);
  return Number.isFinite(Number(String(value).replace(/,/g, "")));
}



function rangeHasSummableCell(sheet, sheetId, range) {
  let found = false;
  forEachCellInRange(range, (row, col) => {
    if (!found && autoSumCellLooksSummable(sheet, sheetId, row, col)) found = true;
  });
  return found;
}



function rangeBottomRowIsBlank(sheet, range) {
  for (let col = range.left; col <= range.right; col += 1) {
    if (cellHasDataForRegion(sheet, range.bottom, col)) return false;
  }
  return true;
}



function rangeRightColumnIsBlank(sheet, range) {
  for (let row = range.top; row <= range.bottom; row += 1) {
    if (cellHasDataForRegion(sheet, row, range.right)) return false;
  }
  return true;
}



function openHomeAutoSumMenu(trigger = null) {
  if (!trigger) {
    applyAutoSumToSelection();
    return;
  }
  openHomeCommandMenu(trigger, "autosum", "オートSUM", `
    ${autoSumMenuButtonHtml("SUM", "合計(S)", "sigma")}
    ${autoSumMenuButtonHtml("AVERAGE", "平均(A)", "bar-chart-3")}
    ${autoSumMenuButtonHtml("COUNT", "数値の個数(C)", "hash")}
    ${autoSumMenuButtonHtml("MAX", "最大値(M)", "chevrons-up")}
    ${autoSumMenuButtonHtml("MIN", "最小値(I)", "chevrons-down")}
    <div class="home-command-menu-separator" role="separator"></div>
    ${autoSumMenuButtonHtml("MORE", "その他の関数(M)...", "function-square")}
  `);
}



function openHomeFillMenu(trigger = null) {
  if (!trigger) {
    fillSelectionFromHome("down");
    return;
  }
  openHomeCommandMenu(trigger, "fill", "フィル", `
    ${fillMenuButtonHtml("down", "下方向へコピー(D)", "arrow-up-down")}
    ${fillMenuButtonHtml("right", "右方向へコピー(R)", "arrow-up-down")}
    ${fillMenuButtonHtml("up", "上方向へコピー(U)", "arrow-up-down")}
    ${fillMenuButtonHtml("left", "左方向へコピー(L)", "arrow-up-down")}
    <div class="home-command-menu-separator" role="separator"></div>
    ${fillMenuButtonHtml("worksheet-fill-dialog", "作業グループへコピー(W)...", "copy-check")}
    ${fillMenuButtonHtml("series-dialog", "連続データ(S)...", "list-plus")}
    ${fillMenuButtonHtml("justify", "両端揃え(J)", "align-justify")}
  `);
}



function openHomeClearMenu(trigger = null) {
  if (!trigger) {
    clearSelectionFromHome("contents");
    return;
  }
  openHomeCommandMenu(trigger, "clear", "クリア", `
    ${clearMenuButtonHtml("all", "すべてクリア(A)", "eraser")}
    ${clearMenuButtonHtml("formats", "書式のクリア(F)", "paintbrush")}
    ${clearMenuButtonHtml("contents", "数式と値のクリア(C)", "delete")}
    <div class="home-command-menu-separator" role="separator"></div>
    ${clearMenuButtonHtml("comments", "コメントとメモのクリア(M)", "message-square-x")}
    ${clearMenuButtonHtml("hyperlinks", "ハイパーリンクのクリア(H)", "unlink")}
  `);
}



function openHomeSortFilterMenu(trigger = null) {
  if (!trigger) {
    openSortDialog();
    return;
  }
  openHomeCommandMenu(trigger, "sort-filter", "並べ替えとフィルター", `
    ${homeCommandMenuButtonHtml("sort-asc", "昇順(S)", "arrow-up-a-z")}
    ${homeCommandMenuButtonHtml("sort-desc", "降順(O)", "arrow-down-z-a")}
    <div class="home-command-menu-separator" role="separator"></div>
    ${homeCommandMenuButtonHtml("sort-dialog", "ユーザー設定の並べ替え(U)...", "list-tree")}
    <div class="home-command-menu-separator" role="separator"></div>
    ${homeCommandMenuButtonHtml("filter", "フィルター(F)", "list-filter")}
    ${homeCommandMenuButtonHtml("clear-filter", "クリア(C)", "eraser")}
    ${homeCommandMenuButtonHtml("reapply-filter", "再適用(R)", "refresh-cw")}
  `);
}



function openFindSelectMenu(trigger = null) {
  if (!trigger) {
    openFindDialog();
    return;
  }
  openHomeCommandMenu(trigger, "find-select", "検索と選択", `
    ${findSelectMenuButtonHtml("find", "検索(F)...", "search")}
    ${findSelectMenuButtonHtml("replace", "置換(R)...", "text-cursor-input")}
    ${findSelectMenuButtonHtml("go-to", "ジャンプ(G)...", "move-right")}
    ${findSelectMenuButtonHtml("go-to-special", "条件を選択(S)...", "list-checks")}
    <div class="home-command-menu-separator" role="separator"></div>
    ${findSelectMenuButtonHtml("select-formulas", "数式(R)", "square-function")}
    ${findSelectMenuButtonHtml("select-comments", "コメント(M)", "message-square-text")}
    ${findSelectMenuButtonHtml("select-notes", "メモ(N)", "sticky-note")}
    ${findSelectMenuButtonHtml("select-conditional-formats", "条件付き書式(C)", "paintbrush")}
    ${findSelectMenuButtonHtml("select-constants", "定数(O)", "hash")}
    ${findSelectMenuButtonHtml("select-validation", "データの入力規則(V)", "check")}
    ${findSelectMenuButtonHtml("select-objects", "オブジェクトの選択(B)", "shapes")}
  `);
}



function openHomeCommandMenu(trigger, kind, label, contentHtml) {
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeRibbonInkEraserMenu();
  closeShapeFormatMenu();
  closeHomeCommandMenu();
  const menu = document.createElement("div");
  menu.id = "homeCommandMenu";
  menu.className = `home-command-menu is-${kind}`;
  menu.dataset.homeCommandKind = kind;
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", label);
  menu.innerHTML = contentHtml;
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  if (window.lucide) window.lucide.createIcons({ icons });
  const rect = trigger.getBoundingClientRect();
  positionPopupMenu(menu, rect.left, rect.bottom + 2);
  state.homeCommandMenu = kind;
  menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
}



function homeCommandMenuButtonHtml(action, label, icon) {
  return homeCommandActionButtonHtml("data-home-sort-filter-action", action, label, icon);
}



function autoSumMenuButtonHtml(action, label, icon) {
  return homeCommandActionButtonHtml("data-home-autosum-action", action, label, icon);
}



function fillMenuButtonHtml(action, label, icon) {
  return homeCommandActionButtonHtml("data-home-fill-action", action, label, icon);
}



function clearMenuButtonHtml(action, label, icon) {
  return homeCommandActionButtonHtml("data-home-clear-action", action, label, icon);
}



function findSelectMenuButtonHtml(action, label, icon) {
  return homeCommandActionButtonHtml("data-find-select-action", action, label, icon);
}



function homeCommandActionButtonHtml(attribute, action, label, icon, options = {}) {
  const accessKey = menuAccessKeyFromLabel(label);
  const accessKeyAttr = accessKey ? ` data-access-key="${escapeAttr(accessKey)}"` : "";
  const role = options.checkable ? "menuitemradio" : "menuitem";
  const checkedAttr = options.checkable ? ` aria-checked="${options.checked ? "true" : "false"}"` : "";
  const disabledAttr = options.disabled ? " disabled" : "";
  const className = [
    options.description ? "has-description" : "",
    options.checked ? "is-checked" : "",
  ].filter(Boolean).join(" ");
  const classAttr = className ? ` class="${className}"` : "";
  const checkHtml = options.checkable ? `<span class="home-command-menu-check" aria-hidden="true">${options.checked ? "✓" : ""}</span>` : "";
  const textHtml = options.description
    ? `<span class="home-command-menu-text"><span class="home-command-menu-label">${accessKeyLabelHtml(label)}</span><small>${escapeHtml(options.description)}</small></span>`
    : `<span>${accessKeyLabelHtml(label)}</span>`;
  return `<button type="button" role="${role}" ${attribute}="${escapeAttr(action)}"${accessKeyAttr}${checkedAttr}${disabledAttr}${classAttr}>${checkHtml}<i data-lucide="${escapeAttr(icon)}"></i>${textHtml}</button>`;
}



function closeHomeCommandMenu() {
  document.querySelector("#homeCommandMenu")?.remove();
  state.homeCommandMenu = null;
}



function closeHomeCommandMenuOnOutsidePointer(event) {
  if (!state.homeCommandMenu) return;
  if (
    event.target.closest?.("#homeCommandMenu") ||
    event.target.closest?.(
      "[data-ribbon-command='autosum-menu'], [data-ribbon-command='fill-menu'], [data-ribbon-command='clear-menu'], [data-ribbon-command='sort-filter'], [data-ribbon-command='find-select'], [data-ribbon-command='formula-autosum'], [data-ribbon-command='formula-function-recent'], [data-ribbon-command='formula-function-financial'], [data-ribbon-command='formula-function-sum'], [data-ribbon-command='formula-function-if'], [data-ribbon-command='formula-function-xlookup'], [data-ribbon-command='formula-function-today'], [data-ribbon-command='formula-function-textjoin'], [data-ribbon-command='formula-function-more'], [data-ribbon-command='formula-define-name'], [data-ribbon-command='formula-use-in-formula'], [data-ribbon-command='layout-themes'], [data-ribbon-command='layout-colors'], [data-ribbon-command='layout-fonts'], [data-ribbon-command='layout-effects'], [data-ribbon-command='layout-margins'], [data-ribbon-command='layout-orientation'], [data-ribbon-command='layout-size'], [data-ribbon-command='layout-print-area'], [data-ribbon-command='layout-breaks']",
    )
  ) {
    return;
  }
  closeHomeCommandMenu();
}



function closeHomeCommandMenuOnEscape(event) {
  if (event.key !== "Escape" || !document.querySelector("#homeCommandMenu")) return;
  if (commandMenuEscapeBlockedByDialog()) return;
  consumeCommandMenuKey(event);
  closeHomeCommandMenu();
}



function commandMenuEscapeBlockedByDialog() {
  return Boolean(document.querySelector("#dataDialog, #formatCellsDialogOverlay, #cellStructureDialogOverlay, #headerSizeDialog, #messageBoxOverlay, #zoomDialog, #findDialog, #goToDialog, #specialDialog, #nameDialogOverlay, #linkDialogOverlay, #objectLinkDialogOverlay, #symbolDialogOverlay, #homeSeriesDialogOverlay, #homeWorksheetFillDialogOverlay"));
}



function enhanceCommandMenuAccessKeys(menu) {
  if (!menu) return;
  menu.querySelectorAll("button").forEach((button) => {
    if (button.dataset.accessKey) return;
    const accessKey = menuAccessKeyFromLabel(button.textContent || button.getAttribute("title") || button.getAttribute("aria-label") || "");
    if (accessKey) button.dataset.accessKey = accessKey;
  });
}



function commandMenuItems(menu) {
  if (!menu) return [];
  return [...menu.querySelectorAll("button")]
    .filter((item) => !item.disabled && !item.classList.contains("is-disabled") && item.offsetParent !== null);
}



function consumeCommandMenuKey(event) {
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
}



function focusCommandMenuItem(items, index) {
  const item = items[index];
  item?.focus?.({ preventScroll: true });
}



function commandMenuItemCenter(item) {
  const rect = item?.getBoundingClientRect?.();
  if (!rect) return { x: 0, y: 0 };
  return {
    x: rect.left + rect.width / 2,
    y: rect.top + rect.height / 2,
    width: rect.width,
    height: rect.height,
  };
}



function nextCommandMenuItemIndexByArrow(items, activeIndex, key) {
  if (!items.length) return -1;
  if (activeIndex < 0) return key === "ArrowUp" || key === "ArrowLeft" ? items.length - 1 : 0;
  const current = commandMenuItemCenter(items[activeIndex]);
  const horizontal = key === "ArrowLeft" || key === "ArrowRight";
  const positive = key === "ArrowRight" || key === "ArrowDown";
  const candidates = items
    .map((item, index) => ({ item, index, center: commandMenuItemCenter(item) }))
    .filter(({ index, center }) => {
      if (index === activeIndex) return false;
      const primaryDelta = horizontal ? center.x - current.x : center.y - current.y;
      return positive ? primaryDelta > 1 : primaryDelta < -1;
    })
    .map((entry) => {
      const primaryDelta = Math.abs((horizontal ? entry.center.x - current.x : entry.center.y - current.y));
      const secondaryDelta = Math.abs((horizontal ? entry.center.y - current.y : entry.center.x - current.x));
      return { ...entry, score: primaryDelta * 1000 + secondaryDelta };
    })
    .sort((left, right) => left.score - right.score || left.index - right.index);
  if (candidates.length) return candidates[0].index;
  return (activeIndex + (positive ? 1 : -1) + items.length) % items.length;
}



function activateCommandMenuItem(item) {
  if (!item || item.disabled || item.classList.contains("is-disabled")) return;
  item.click();
  restoreWorksheetFocusAfterCommandPopup();
}



function closeActiveCommandMenu(menu) {
  if (!menu) return false;
  switch (menu.id) {
    case "homeCommandMenu":
      closeHomeCommandMenu();
      return true;
    case "dataCommandMenu":
      closeDataCommandMenu();
      return true;
    case "dataValidationMenu":
      closeDataValidationMenu();
      return true;
    case "whatIfMenu":
      closeWhatIfMenu();
      return true;
    case "viewCommandMenu":
      closeViewCommandMenu();
      return true;
    case "fileSideMenu":
      closeFileSideMenu({ restoreHome: true });
      restoreWorksheetFocusAfterCommandPopup();
      return true;
    case "nameBoxMenu":
      closeNameBoxMenu({ restoreFocus: true });
      return true;
    case "ribbonUnderlineMenu":
      closeRibbonUnderlineMenu();
      return true;
    case "ribbonBorderMenu":
      closeRibbonBorderMenu();
      return true;
    case "ribbonInkEraserMenu":
      closeRibbonInkEraserMenu();
      return true;
    case "ribbonColorPalette":
      closeRibbonColorPalette();
      return true;
    case "ribbonShapeMenu":
      closeRibbonShapeMenu();
      return true;
    case "shapeFormatMenu":
      closeShapeFormatMenu();
      return true;
    case "sheetTabContextMenu":
      closeSheetTabContextMenu();
      return true;
    case "sheetTabNavigatorMenu":
      closeSheetNavigatorMenu();
      return true;
    case "cellContextMenu":
      closeCellContextMenu();
      return true;
    case "statusBarContextMenu":
      closeStatusBarContextMenu();
      return true;
    case "queryPaneContextMenu":
      closeQueryPaneContextMenu();
      return true;
    case "quickAnalysisPopup":
      closeQuickAnalysisPopup();
      return true;
    default:
      return false;
  }
}



function handleRibbonCommandMenuKeydown(event) {
  if (isElementImeCompositionEvent(event) || event.defaultPrevented) return;
  const menu = event.currentTarget;
  if (event.key === "Escape") {
    if (closeActiveCommandMenu(menu)) consumeCommandMenuKey(event);
    return;
  }
  const items = commandMenuItems(menu);
  if (!items.length) return;
  const activeButton = document.activeElement?.closest?.("button");
  const activeIndex = items.includes(activeButton) ? items.indexOf(activeButton) : -1;
  if (["ArrowDown", "ArrowUp", "ArrowLeft", "ArrowRight"].includes(event.key)) {
    focusCommandMenuItem(items, nextCommandMenuItemIndexByArrow(items, activeIndex, event.key));
    consumeCommandMenuKey(event);
    return;
  }
  if (event.key === "Home" || event.key === "End") {
    focusCommandMenuItem(items, event.key === "Home" ? 0 : items.length - 1);
    consumeCommandMenuKey(event);
    return;
  }
  if (event.key === "Enter" || event.key === " ") {
    const target = activeButton && items.includes(activeButton) ? activeButton : items[0];
    consumeCommandMenuKey(event);
    activateCommandMenuItem(target);
    return;
  }
  const accessKey = accessKeyFromKeyboardEvent(event);
  if (accessKey && !event.ctrlKey && !event.metaKey && (event.altKey || event.key?.length === 1)) {
    const target = items.find((item) => String(item.dataset.accessKey || menuAccessKeyFromLabel(item.textContent || "")).toUpperCase() === accessKey);
    if (!target) return;
    consumeCommandMenuKey(event);
    activateCommandMenuItem(target);
  }
}



function activeCommandMenuForGlobalAccessKey() {
  return document.querySelector("#homeCommandMenu, #dataCommandMenu, #dataValidationMenu, #whatIfMenu, #viewCommandMenu, #ribbonUnderlineMenu, #ribbonBorderMenu, #ribbonInkEraserMenu, #ribbonColorPalette, #ribbonShapeMenu, #shapeFormatMenu, #sheetTabContextMenu, #sheetTabNavigatorMenu, #cellContextMenu, #statusBarContextMenu, #queryPaneContextMenu, #quickAnalysisPopup, #fileSideMenu:not(.hidden), #nameBoxMenu");
}



function handleGlobalCommandMenuAccessKey(event) {
  if (isElementImeCompositionEvent(event) || event.defaultPrevented) return;
  const menu = activeCommandMenuForGlobalAccessKey();
  if (!menu) return;
  const accessKey = accessKeyFromKeyboardEvent(event);
  if (!accessKey || event.ctrlKey || event.metaKey || !(event.altKey || event.key?.length === 1)) return;
  const target = commandMenuItems(menu).find((item) =>
    String(item.dataset.accessKey || menuAccessKeyFromLabel(item.textContent || "")).toUpperCase() === accessKey,
  );
  if (!target) return;
  consumeCommandMenuKey(event);
  activateCommandMenuItem(target);
}



function handleHomeSortFilterMenuAction(event) {
  const action = event.currentTarget.dataset.homeSortFilterAction;
  closeHomeCommandMenu();
  if (action === "sort-asc") {
    sortSelectionByActiveColumn("asc");
  } else if (action === "sort-desc") {
    sortSelectionByActiveColumn("desc");
  } else if (action === "sort-dialog") {
    openSortDialog();
  } else if (action === "filter") {
    toggleAutoFilterForSelection();
  } else if (action === "clear-filter") {
    clearAutoFilterCriteria();
  } else if (action === "reapply-filter") {
    reapplyAutoFilter();
  }
}



function handleHomeAutoSumMenuAction(event) {
  const action = event.currentTarget.dataset.homeAutosumAction;
  closeHomeCommandMenu();
  if (action === "MORE") {
    openInsertFunctionDialog({ category: "すべて表示" });
    return;
  }
  applyAutoSumToSelection(action);
}



function handleFormulaCalculationMenuAction(event) {
  const action = event.currentTarget.dataset.formulaCalculationAction || "automatic";
  closeHomeCommandMenu();
  setCalculationMode(action);
}



function handleHomeFillMenuAction(event) {
  const action = event.currentTarget.dataset.homeFillAction;
  closeHomeCommandMenu();
  if (action === "series-dialog") {
    openHomeSeriesDialog();
    return;
  }
  if (action === "justify") {
    applyHomeJustifyFill();
    return;
  }
  if (action === "worksheet-fill-dialog") {
    openHomeWorksheetFillDialog();
    return;
  }
  fillSelectionFromHome(action);
}



function openHomeSeriesDialog() {
  closeHomeSeriesDialog({ restoreFocus: false });
  const range = activeSelectionRanges()[0] || activeSelectionRange();
  const preferRows = range && range.right - range.left > range.bottom - range.top;
  const overlay = document.createElement("div");
  overlay.id = "homeSeriesDialogOverlay";
  overlay.className = "excel-dialog-overlay";
  overlay.previousFocus = document.activeElement;
  overlay.innerHTML = `
    <div id="homeSeriesDialog" class="excel-mini-dialog data-dialog home-series-dialog" role="dialog" aria-modal="true" aria-labelledby="homeSeriesDialogTitle">
      <div class="excel-mini-titlebar" id="homeSeriesDialogTitle">連続データ</div>
      <div class="excel-mini-body data-dialog-body home-series-dialog-body">
        <fieldset class="data-dialog-fieldset home-series-fieldset">
          <legend>範囲</legend>
          <label class="data-dialog-check"><input type="radio" name="homeSeriesDirection" value="rows"${preferRows ? " checked" : ""} data-home-series-field />${accessKeyLabelHtml("行(R)")}</label>
          <label class="data-dialog-check"><input type="radio" name="homeSeriesDirection" value="columns"${preferRows ? "" : " checked"} data-home-series-field />${accessKeyLabelHtml("列(C)")}</label>
        </fieldset>
        <fieldset class="data-dialog-fieldset home-series-fieldset">
          <legend>種類</legend>
          <label class="data-dialog-check"><input type="radio" name="homeSeriesType" value="linear" checked data-home-series-field />${accessKeyLabelHtml("加算(L)")}</label>
          <label class="data-dialog-check"><input type="radio" name="homeSeriesType" value="growth" data-home-series-field />${accessKeyLabelHtml("乗算(G)")}</label>
          <label class="data-dialog-check"><input type="radio" name="homeSeriesType" value="date" data-home-series-field />${accessKeyLabelHtml("日付(D)")}</label>
          <label class="data-dialog-check"><input type="radio" name="homeSeriesType" value="auto" data-home-series-field />${accessKeyLabelHtml("オートフィル(A)")}</label>
        </fieldset>
        <fieldset class="data-dialog-fieldset home-series-fieldset" data-home-series-date-unit>
          <legend>日付の単位</legend>
          <label class="data-dialog-check"><input type="radio" name="homeSeriesDateUnit" value="day" checked data-home-series-field />${accessKeyLabelHtml("日(D)")}</label>
          <label class="data-dialog-check"><input type="radio" name="homeSeriesDateUnit" value="weekday" data-home-series-field />${accessKeyLabelHtml("週日(W)")}</label>
          <label class="data-dialog-check"><input type="radio" name="homeSeriesDateUnit" value="month" data-home-series-field />${accessKeyLabelHtml("月(M)")}</label>
          <label class="data-dialog-check"><input type="radio" name="homeSeriesDateUnit" value="year" data-home-series-field />${accessKeyLabelHtml("年(Y)")}</label>
        </fieldset>
        <div class="home-series-input-grid">
          <label>${accessKeyLabelHtml("増分値(S):")}<input type="text" value="1" data-home-series-step data-home-series-field /></label>
          <label>${accessKeyLabelHtml("停止値(O):")}<input type="text" value="" data-home-series-stop data-home-series-field /></label>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-home-series-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-home-series-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.append(overlay);
  enhanceDialogAccessKeys(overlay);
  syncHomeSeriesDialog(overlay);
  overlay.querySelector("[data-home-series-step]")?.focus({ preventScroll: true });
}



function closeHomeSeriesDialog(options = {}) {
  const previousFocus = document.querySelector("#homeSeriesDialogOverlay")?.previousFocus || null;
  document.querySelector("#homeSeriesDialogOverlay")?.remove();
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) previousFocus.focus?.({ preventScroll: true });
}



function handleHomeSeriesDialogAction(event) {
  const action = event.currentTarget.dataset.homeSeriesDialogAction || "cancel";
  if (action === "cancel") {
    closeHomeSeriesDialog();
    return;
  }
  applyHomeSeriesDialog();
}



function handleHomeSeriesDialogKeydown(event) {
  if (isElementImeCompositionEvent(event)) return;
  const accessKey = accessKeyFromKeyboardEvent(event);
  if (accessKey && event.altKey && !event.ctrlKey && !event.metaKey) {
    const target = findDialogAccessKeyTarget(event.currentTarget, accessKey);
    if (target) {
      event.preventDefault();
      activateDialogAccessKeyTarget(target);
    }
    return;
  }
  if (event.key === "Enter") {
    event.preventDefault();
    applyHomeSeriesDialog();
  } else if (event.key === "Escape") {
    event.preventDefault();
    closeHomeSeriesDialog();
  }
}



function handleHomeSeriesDialogSync(event) {
  syncHomeSeriesDialog(event.currentTarget.closest("#homeSeriesDialogOverlay"));
}



function syncHomeSeriesDialog(overlay = document.querySelector("#homeSeriesDialogOverlay")) {
  if (!overlay) return;
  const isDate = overlay.querySelector('input[name="homeSeriesType"]:checked')?.value === "date";
  overlay.querySelectorAll("[data-home-series-date-unit] input").forEach((input) => {
    input.disabled = !isDate;
  });
}



function collectHomeSeriesDialogValues(dialog = document.querySelector("#homeSeriesDialog")) {
  const type = dialog?.querySelector('input[name="homeSeriesType"]:checked')?.value || "linear";
  return {
    direction: dialog?.querySelector('input[name="homeSeriesDirection"]:checked')?.value || "columns",
    type,
    dateUnit: dialog?.querySelector('input[name="homeSeriesDateUnit"]:checked')?.value || "day",
    step: Number.parseFloat(String(dialog?.querySelector("[data-home-series-step]")?.value || "1").replace(",", ".")),
    stop: String(dialog?.querySelector("[data-home-series-stop]")?.value || "").trim(),
  };
}



function applyHomeSeriesDialog() {
  const dialog = document.querySelector("#homeSeriesDialog");
  if (!dialog || !state.model) return;
  const values = collectHomeSeriesDialogValues(dialog);
  const range = activeSelectionRanges()[0];
  if (!range || range.fullRow || range.fullColumn) {
    showMessageBox("連続データを作成するセル範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const step = Number.isFinite(values.step) && values.step !== 0 ? values.step : 1;
  const sheet = activeSheet();
  if (!ensureEditableRangeForProtection(sheet, range, "連続データ")) return;
  const before = captureCellsForHistory(sheet, range);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = rawActiveSelectionRanges().map((item) => ({ ...item }));
  let wrote = false;
  ensureSheetSize(sheet, range.bottom, range.right);
  if (values.direction === "rows") {
    for (let row = range.top; row <= range.bottom; row += 1) {
      wrote = writeHomeSeriesLine(sheet, range, [{ row, col: range.left }], { axis: "col", type: values.type, dateUnit: values.dateUnit, step, stop: values.stop }) || wrote;
    }
  } else {
    for (let col = range.left; col <= range.right; col += 1) {
      wrote = writeHomeSeriesLine(sheet, range, [{ row: range.top, col }], { axis: "row", type: values.type, dateUnit: values.dateUnit, step, stop: values.stop }) || wrote;
    }
  }
  if (!wrote) {
    showMessageBox("連続データを作成する開始値がありません。", { title: "Cht WebSheet" });
    return;
  }
  internCellsInRange(sheet, range);
  state.hf = buildFormulaEngine(state.model);
  const after = captureCellsForHistory(sheet, range);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore,
    selectionAfter: range,
    selectionBeforeRanges,
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} 連続データ`,
  });
  closeHomeSeriesDialog({ restoreFocus: false });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${rangeToLabel(range)} に連続データを作成しました。`);
}



function writeHomeSeriesLine(sheet, range, seeds, options) {
  const seed = seeds[0];
  const seedCell = sheet.cells[cellKey(seed.row, seed.col)];
  const seedValue = homeSeriesSeedValue(seedCell);
  if (seedValue == null) return false;
  const maxOffset = options.axis === "row" ? range.bottom - range.top : range.right - range.left;
  for (let offset = 0; offset <= maxOffset; offset += 1) {
    const row = options.axis === "row" ? range.top + offset : seed.row;
    const col = options.axis === "col" ? range.left + offset : seed.col;
    const nextValue = homeSeriesValueForOffset(seedValue, offset, options);
    if (nextValue == null) continue;
    writeHomeSeriesCell(sheet, row, col, seedCell, nextValue);
  }
  return true;
}



function homeSeriesSeedValue(cell) {
  if (!cell) return null;
  const raw = cell.formula ? cell.cached : cell.raw ?? cell.cached;
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) return dateToExcelSerial(raw);
  if (typeof raw === "number" && Number.isFinite(raw)) return raw;
  const numeric = Number.parseFloat(String(raw ?? "").replace(",", "."));
  return Number.isFinite(numeric) ? numeric : null;
}



function homeSeriesValueForOffset(seedValue, offset, options) {
  if (options.type === "growth") return seedValue * Math.pow(options.step, offset);
  if (options.type === "date") return homeSeriesDateValue(seedValue, offset, options.step, options.dateUnit);
  return seedValue + options.step * offset;
}



function homeSeriesDateValue(seedValue, offset, step, unit) {
  if (unit === "weekday") {
    let serial = seedValue;
    let remaining = offset * step;
    const direction = remaining < 0 ? -1 : 1;
    remaining = Math.abs(remaining);
    while (remaining > 0) {
      serial += direction;
      const day = excelSerialToDate(serial).getDay();
      if (day !== 0 && day !== 6) remaining -= 1;
    }
    return serial;
  }
  if (unit === "month" || unit === "year") {
    const date = excelSerialToDate(seedValue);
    if (unit === "month") date.setMonth(date.getMonth() + offset * step);
    else date.setFullYear(date.getFullYear() + offset * step);
    return dateToExcelSerial(date);
  }
  return seedValue + offset * step;
}



function writeHomeSeriesCell(sheet, row, col, sourceCell, value) {
  const key = cellKey(row, col);
  const nextCell = cloneCellModel(sourceCell) || ensureCellModel(sheet, row, col);
  nextCell.row = row;
  nextCell.col = col;
  nextCell.raw = value;
  nextCell.formula = null;
  nextCell.cached = value;
  nextCell.display = displayValue(value, nextCell.numFmt);
  nextCell.richText = null;
  storeCellModelForSheet(sheet, key, nextCell);
}



function applyHomeJustifyFill() {
  const ranges = activeSelectionRanges();
  if (ranges.length !== 1) {
    showMessageBox("両端揃えを実行する範囲を1つ選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const range = ranges[0];
  if (range.fullRow || range.fullColumn) {
    showMessageBox("行全体または列全体の選択範囲には、両端揃えを実行できません。", { title: "Cht WebSheet" });
    return;
  }
  const sheet = activeSheet();
  if (!ensureEditableRangeForProtection(sheet, range, "両端揃え")) return;
  const textCells = [];
  forEachCellInRange(range, (row, col) => {
    const text = String(getCellRawInput(sheet, row, col) || "").trim();
    if (text) textCells.push({ row, col, text, cell: sheet.cells[cellKey(row, col)] });
  });
  if (!textCells.length) {
    showMessageBox("両端揃えする文字列がありません。", { title: "Cht WebSheet" });
    return;
  }
  const before = captureCellsForHistory(sheet, range);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = rawActiveSelectionRanges().map((item) => ({ ...item }));
  const words = textCells.map((item) => item.text).join(" ").split(/\s+/).filter(Boolean);
  const targets = homeJustifyTargets(range);
  const template = textCells[0].cell || null;
  const chunks = homeJustifyTextChunks(words, targets.length, Math.max(6, Math.floor(homeJustifyRangePixelWidth(sheet, range, targets) / 8)));
  targets.forEach((target, index) => {
    const key = cellKey(target.row, target.col);
    const value = chunks[index] || "";
    if (!value) {
      delete sheet.cells[key];
      return;
    }
    const nextCell = cloneCellModel(template) || ensureCellModel(sheet, target.row, target.col);
    nextCell.row = target.row;
    nextCell.col = target.col;
    nextCell.raw = value;
    nextCell.formula = null;
    nextCell.cached = value;
    nextCell.display = value;
    nextCell.richText = null;
    storeCellModelForSheet(sheet, key, nextCell);
  });
  internCellsInRange(sheet, range);
  state.hf = buildFormulaEngine(state.model);
  const after = captureCellsForHistory(sheet, range);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore,
    selectionAfter: range,
    selectionBeforeRanges,
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} 両端揃え`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${rangeToLabel(range)} に両端揃えを適用しました。`);
}



function homeJustifyTargets(range) {
  const targets = [];
  if (range.bottom - range.top >= range.right - range.left) {
    for (let row = range.top; row <= range.bottom; row += 1) targets.push({ row, col: range.left });
  } else {
    for (let col = range.left; col <= range.right; col += 1) targets.push({ row: range.top, col });
  }
  return targets;
}



function homeJustifyRangePixelWidth(sheet, range, targets) {
  if (!targets.length) return 64;
  if (targets[0].row === range.top && targets.length > 1) {
    return targets.reduce((total, target) => total + columnWidthForIndex(sheet, target.col), 0);
  }
  return columnWidthForIndex(sheet, targets[0].col);
}



function homeJustifyTextChunks(words, targetCount, maxChars) {
  const chunks = Array.from({ length: targetCount }, () => "");
  let index = 0;
  words.forEach((word) => {
    const current = chunks[index] || "";
    if (current && `${current} ${word}`.length > maxChars && index < chunks.length - 1) index += 1;
    chunks[index] = chunks[index] ? `${chunks[index]} ${word}` : word;
  });
  return chunks;
}



function columnWidthForIndex(sheet, col) {
  const column = sheet.columns?.[col - 1];
  return Number(column?.width || sheet.defaultColumnWidth || columnWidthToPixels(8.43));
}



function openHomeWorksheetFillDialog() {
  closeHomeWorksheetFillDialog({ restoreFocus: false });
  const grouped = activeSheetGroupEditIndexes();
  if (grouped.length <= 1) {
    showMessageBox("作業グループへコピーするには、複数のシートを選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const overlay = document.createElement("div");
  overlay.id = "homeWorksheetFillDialogOverlay";
  overlay.className = "excel-dialog-overlay";
  overlay.previousFocus = document.activeElement;
  overlay.innerHTML = `
    <div id="homeWorksheetFillDialog" class="excel-mini-dialog data-dialog home-worksheet-fill-dialog" role="dialog" aria-modal="true" aria-labelledby="homeWorksheetFillDialogTitle">
      <div class="excel-mini-titlebar" id="homeWorksheetFillDialogTitle">作業グループへコピー</div>
      <div class="excel-mini-body data-dialog-body">
        <fieldset class="data-dialog-fieldset home-series-fieldset">
          <legend>コピーする内容</legend>
          <label class="data-dialog-check"><input type="radio" name="homeWorksheetFillMode" value="all" checked />${accessKeyLabelHtml("すべて(A)")}</label>
          <label class="data-dialog-check"><input type="radio" name="homeWorksheetFillMode" value="contents" />${accessKeyLabelHtml("数式と値(C)")}</label>
          <label class="data-dialog-check"><input type="radio" name="homeWorksheetFillMode" value="formats" />${accessKeyLabelHtml("書式(F)")}</label>
        </fieldset>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-home-worksheet-fill-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-home-worksheet-fill-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.append(overlay);
  enhanceDialogAccessKeys(overlay);
  overlay.querySelector("input")?.focus({ preventScroll: true });
}



function closeHomeWorksheetFillDialog(options = {}) {
  const overlay = document.querySelector("#homeWorksheetFillDialogOverlay");
  const previousFocus = overlay?.previousFocus || null;
  overlay?.remove();
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) previousFocus.focus?.({ preventScroll: true });
  else if (options.restoreFocus !== false) restoreWorksheetFocusAfterCommandPopup();
}



function handleHomeWorksheetFillDialogAction(event) {
  const action = event.currentTarget.dataset.homeWorksheetFillAction || "cancel";
  if (action === "cancel") {
    closeHomeWorksheetFillDialog();
    return;
  }
  applyHomeWorksheetFillDialog();
}



function handleHomeWorksheetFillDialogKeydown(event) {
  if (isElementImeCompositionEvent(event)) return;
  const accessKey = accessKeyFromKeyboardEvent(event);
  if (accessKey && event.altKey && !event.ctrlKey && !event.metaKey) {
    const target = findDialogAccessKeyTarget(event.currentTarget, accessKey);
    if (target) {
      event.preventDefault();
      activateDialogAccessKeyTarget(target);
    }
    return;
  }
  if (event.key === "Enter") {
    event.preventDefault();
    applyHomeWorksheetFillDialog();
  } else if (event.key === "Escape") {
    event.preventDefault();
    closeHomeWorksheetFillDialog();
  }
}



function applyHomeWorksheetFillDialog() {
  const dialog = document.querySelector("#homeWorksheetFillDialog");
  if (!dialog) return;
  const mode = dialog.querySelector('input[name="homeWorksheetFillMode"]:checked')?.value || "all";
  const ranges = activeSelectionRanges();
  if (!ranges.length) {
    showMessageBox("コピーするセル範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const sheetIndexes = activeSheetGroupEditIndexes().filter((index) => index !== state.activeSheetIndex);
  const sourceSheet = activeSheet();
  const items = [];
  sheetIndexes.forEach((sheetIndex) => {
    const targetSheet = state.model.sheets[sheetIndex];
    if (!targetSheet) return;
    const targetRanges = materializedRangesForSheet(targetSheet, ranges, sheetIndex);
    if (!targetRanges.length) return;
    targetRanges.forEach((range) => {
      if (ensureEditableRangeForProtection(targetSheet, range, "作業グループへコピー")) items.push({ sheetIndex, sheet: targetSheet, range });
    });
  });
  if (!items.length) {
    showMessageBox("コピー先のシートがありません。", { title: "Cht WebSheet" });
    return;
  }
  const sheetsBefore = captureSheetsForHistory([state.activeSheetIndex, ...sheetIndexes]);
  items.forEach((item) => {
    ensureSheetSize(item.sheet, item.range.bottom, item.range.right);
    forEachCellInRange(item.range, (row, col) => {
      const sourceCell = sourceSheet.cells[cellKey(row, col)] || null;
      copyWorksheetFillCell(item.sheet, row, col, sourceCell, mode);
    });
    internCellsInRange(item.sheet, item.range);
  });
  state.hf = buildFormulaEngine(state.model);
  closeHomeWorksheetFillDialog({ restoreFocus: false });
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: rawActiveSelectionRanges().map((item) => ({ ...item })),
    selectionAfterRanges: rawActiveSelectionRanges().map((item) => ({ ...item })),
    label: `${items.length} シート 作業グループへコピー`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${items.length} 個のシートへコピーしました。`);
}



function copyWorksheetFillCell(targetSheet, row, col, sourceCell, mode) {
  const key = cellKey(row, col);
  const targetCell = targetSheet.cells[key] || null;
  if (mode === "all") {
    if (!sourceCell) delete targetSheet.cells[key];
    else {
      const next = cloneCellModel(sourceCell);
      next.row = row;
      next.col = col;
      storeCellModelForSheet(targetSheet, key, next);
    }
    return;
  }
  const next = cloneCellModel(targetCell) || ensureCellModel(targetSheet, row, col);
  if (mode === "contents") {
    next.raw = cloneCellScalarValue(sourceCell?.raw);
    next.formula = sourceCell?.formula || null;
    next.cached = cloneCellScalarValue(sourceCell?.cached);
    next.display = sourceCell?.display || "";
    next.hyperlink = cloneJsonLikeValue(sourceCell?.hyperlink);
    next.richText = cloneJsonLikeValue(sourceCell?.richText);
  } else if (mode === "formats" && sourceCell) {
    next.css = sourceCell.css || {};
    next.hasFill = Boolean(sourceCell.hasFill);
    next.hasFontColor = Boolean(sourceCell.hasFontColor);
    next.borders = sourceCell.borders || {};
    next.numFmt = sourceCell.numFmt || "";
    next.protection = cloneJsonLikeValue(sourceCell.protection);
  }
  next.row = row;
  next.col = col;
  storeCellModelForSheet(targetSheet, key, next);
}



function handleHomeClearMenuAction(event) {
  const action = event.currentTarget.dataset.homeClearAction;
  closeHomeCommandMenu();
  clearSelectionFromHome(action);
}



function fillSelectionFromHome(direction = "down") {
  const fill = homeFillRanges(direction);
  if (!fill) return;
  if (fillSelectionFromHomeSheetGroup(fill)) return;
  const { sheet, range, sourceRange, fillRange, fillDirection, selectionBefore, selectionBeforeRanges } = fill;
  if (!ensureEditableRangeForProtection(sheet, fillRange, "オートフィル")) return;
  const mergeValidation = validateAutoFillMerges(sheet, sourceRange, fillRange, fillDirection);
  if (!mergeValidation.ok) {
    showMessageBox(mergeValidation.message, { title: "Cht WebSheet" });
    setStatus(mergeValidation.message);
    return;
  }

  const before = captureCellsForHistory(sheet, range);
  const mergesBefore = cloneMerges(sheet.merges);
  ensureSheetSize(sheet, range.bottom, range.right);
  removeIntersectingMerges(sheet, fillRange);
  writeHomeFillCells(sheet, sourceRange, fillRange, fillDirection);
  applyAutoFillMerges(sheet, sourceRange, fillRange, fillDirection);
  state.hf = buildFormulaEngine(state.model);
  const after = captureCellsForHistory(sheet, range);
  const mergesAfter = cloneMerges(sheet.merges);
  state.selected = { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    mergesBefore,
    mergesAfter,
    selectionBefore,
    selectionAfter: range,
    selectionBeforeRanges,
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} フィル`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${rangeToLabel(range)} にフィルを適用しました。`);
}



function homeFillRanges(direction = "down") {
  const fillDirection = ["down", "right", "up", "left"].includes(direction) ? direction : "down";
  const ranges = activeSelectionRanges();
  if (ranges.length !== 1) {
    showMessageBox("フィルを実行する範囲を1つ選択してください。", { title: "Cht WebSheet" });
    return null;
  }
  const range = ranges[0];
  if (range.fullRow || range.fullColumn) {
    showMessageBox("行全体または列全体の選択範囲には、このフィルを実行できません。", { title: "Cht WebSheet" });
    return null;
  }
  const height = range.bottom - range.top + 1;
  const width = range.right - range.left + 1;
  if ((fillDirection === "down" || fillDirection === "up") && height < 2) {
    showMessageBox("コピー元とコピー先を含む2行以上の範囲を選択してください。", { title: "Cht WebSheet" });
    return null;
  }
  if ((fillDirection === "right" || fillDirection === "left") && width < 2) {
    showMessageBox("コピー元とコピー先を含む2列以上の範囲を選択してください。", { title: "Cht WebSheet" });
    return null;
  }
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = rawActiveSelectionRanges().map((item) => ({ ...item }));
  const sheet = activeSheet();
  const plan = homeFillPlanForRange(sheet, range, fillDirection, state.activeSheetIndex);
  return { ...plan, selectionBefore, selectionBeforeRanges };
}



function homeFillPlanForRange(sheet, range, fillDirection, sheetIndex = state.activeSheetIndex) {
  let sourceRange = null;
  let fillRange = null;
  if (fillDirection === "down") {
    sourceRange = normalizeRange({ sheetIndex, row: range.top, col: range.left }, { sheetIndex, row: range.top, col: range.right });
    fillRange = normalizeRange({ sheetIndex, row: range.top + 1, col: range.left }, { sheetIndex, row: range.bottom, col: range.right });
  } else if (fillDirection === "up") {
    sourceRange = normalizeRange({ sheetIndex, row: range.bottom, col: range.left }, { sheetIndex, row: range.bottom, col: range.right });
    fillRange = normalizeRange({ sheetIndex, row: range.top, col: range.left }, { sheetIndex, row: range.bottom - 1, col: range.right });
  } else if (fillDirection === "right") {
    sourceRange = normalizeRange({ sheetIndex, row: range.top, col: range.left }, { sheetIndex, row: range.bottom, col: range.left });
    fillRange = normalizeRange({ sheetIndex, row: range.top, col: range.left + 1 }, { sheetIndex, row: range.bottom, col: range.right });
  } else {
    sourceRange = normalizeRange({ sheetIndex, row: range.top, col: range.right }, { sheetIndex, row: range.bottom, col: range.right });
    fillRange = normalizeRange({ sheetIndex, row: range.top, col: range.left }, { sheetIndex, row: range.bottom, col: range.right - 1 });
  }
  return { sheet, range: { ...range, sheetIndex }, sourceRange, fillRange, fillDirection };
}



function fillSelectionFromHomeSheetGroup(fill) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !fill?.range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const targetRanges = materializedRangesForSheet(sheet, [fill.range], sheetIndex);
    const targetRange = targetRanges[0];
    if (!targetRange) continue;
    const plan = homeFillPlanForRange(sheet, targetRange, fill.fillDirection, sheetIndex);
    if (!ensureEditableRangeForProtection(sheet, plan.fillRange, "オートフィル")) return true;
    const mergeValidation = validateAutoFillMerges(sheet, plan.sourceRange, plan.fillRange, plan.fillDirection);
    if (!mergeValidation.ok) {
      const message = `${sheet.name}!${rangeToLabel(plan.range)}: ${mergeValidation.message}`;
      showMessageBox(message, { title: "Cht WebSheet" });
      setStatus(mergeValidation.message);
      return true;
    }
    items.push({ sheetIndex, sheet, ...plan, ranges: [plan.range] });
  }
  if (!items.length) return true;
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  try {
    items.forEach((item) => {
      ensureSheetSize(item.sheet, item.range.bottom, item.range.right);
      removeIntersectingMerges(item.sheet, item.fillRange);
      writeHomeFillCells(item.sheet, item.sourceRange, item.fillRange, item.fillDirection);
      applyAutoFillMerges(item.sheet, item.sourceRange, item.fillRange, item.fillDirection);
      internCellsInRange(item.sheet, item.range);
    });
    refreshChartsForGroupedRanges(items);
    state.hf = buildFormulaEngine(state.model);
    state.selected = { sheetIndex: state.activeSheetIndex, row: fill.range.top, col: fill.range.left };
    state.selectionRange = fill.range;
    state.selectionRanges = [fill.range];
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.selectedImages = [];
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: fill.selectionBefore,
      selectionAfter: fill.range,
      selectionBeforeRanges: fill.selectionBeforeRanges,
      selectionAfterRanges: [fill.range],
      label: `${rangeToLabel(fill.range)} ${sheetIndexes.length} シートフィル`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangeToLabel(fill.range)} にフィルを適用しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`フィルを適用できませんでした：${error.message || error}`);
  }
  return true;
}



function writeHomeFillCells(sheet, sourceRange, fillRange, direction) {
  forEachCellInRange(fillRange, (row, col) => {
    const sourcePoint = homeFillSourcePoint(sourceRange, row, col, direction);
    const sourceCell = sheet.cells[cellKey(sourcePoint.row, sourcePoint.col)] || null;
    const key = cellKey(row, col);
    if (!sourceCell) {
      delete sheet.cells[key];
      return;
    }
    const nextCell = cloneCellModel(sourceCell);
    const rowOffset = row - sourcePoint.row;
    const colOffset = col - sourcePoint.col;
    if (typeof nextCell.formula === "string" && nextCell.formula.startsWith("=")) {
      nextCell.formula = shiftFormulaReferences(nextCell.formula, rowOffset, colOffset);
      nextCell.raw = nextCell.formula;
      nextCell.cached = null;
      nextCell.display = nextCell.formula;
    }
    nextCell.row = row;
    nextCell.col = col;
    storeCellModelForSheet(sheet, key, nextCell);
  });
}



function homeFillSourcePoint(sourceRange, row, col, direction) {
  if (direction === "down" || direction === "up") {
    return { row: direction === "down" ? sourceRange.top : sourceRange.bottom, col };
  }
  return { row, col: direction === "right" ? sourceRange.left : sourceRange.right };
}



function clearSelectionFromHome(mode = "contents") {
  const normalizedMode = ["all", "formats", "contents", "comments", "hyperlinks"].includes(mode) ? mode : "contents";
  const ranges = activeSelectionRanges();
  if (!ranges.length) {
    showMessageBox("クリアするセルを選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const cellCount = ranges.reduce((total, range) => total + (range.bottom - range.top + 1) * (range.right - range.left + 1), 0);
  if (cellCount > DATA_OPERATION_CELL_LIMIT) {
    showMessageBox("選択範囲が大きすぎるため、クリアできません。", { title: "Cht WebSheet" });
    return;
  }

  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = rawActiveSelectionRanges().map((range) => ({ ...range }));
  if (clearSelectionFromHomeSheetGroup(normalizedMode, ranges, selectionBefore, selectionBeforeRanges)) return;

  const sheet = activeSheet();
  const before = captureCellsForHistoryRanges(sheet, ranges);
  const seen = new Set();
  ranges.forEach((range) => {
    forEachCellInRange(range, (row, col) => {
      const key = cellKey(row, col);
      if (seen.has(key)) return;
      seen.add(key);
      clearCellModelForHome(sheet, key, normalizedMode);
    });
  });
  state.hf = buildFormulaEngine(state.model);
  const after = captureCellsForHistoryRanges(sheet, ranges);
  const firstRange = ranges[0];
  state.selected = { sheetIndex: state.activeSheetIndex, row: firstRange.top, col: firstRange.left };
  state.selectionRange = firstRange;
  state.selectionRanges = ranges.map((range) => ({ ...range }));
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore,
    selectionAfter: firstRange,
    selectionBeforeRanges,
    selectionAfterRanges: ranges.map((range) => ({ ...range })),
    label: `${rangesToLabel(ranges)} ${homeClearLabel(normalizedMode)}`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${rangesToLabel(ranges)} の${homeClearLabel(normalizedMode)}を実行しました。`);
}



function clearSelectionFromHomeSheetGroup(normalizedMode, ranges, selectionBefore, selectionBeforeRanges) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !ranges?.length) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const targetRanges = materializedRangesForSheet(sheet, ranges, sheetIndex);
    if (!targetRanges.length) continue;
    if (!ensureEditableRangesForProtection(sheet, targetRanges, homeClearLabel(normalizedMode))) return true;
    items.push({ sheetIndex, sheet, ranges: targetRanges });
  }
  if (!items.length) return true;
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  try {
    items.forEach((item) => {
      const seen = new Set();
      item.ranges.forEach((range) => {
        forEachCellInRange(range, (row, col) => {
          const key = cellKey(row, col);
          if (seen.has(key)) return;
          seen.add(key);
          clearCellModelForHome(item.sheet, key, normalizedMode);
        });
      });
      internCellsInRanges(item.sheet, item.ranges);
    });
    refreshChartsForGroupedRanges(items);
    state.hf = buildFormulaEngine(state.model);
    const firstRange = ranges[0];
    state.selected = { sheetIndex: state.activeSheetIndex, row: firstRange.top, col: firstRange.left };
    state.selectionRange = firstRange;
    state.selectionRanges = ranges.map((range) => ({ ...range }));
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.selectedImages = [];
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore,
      selectionAfter: firstRange,
      selectionBeforeRanges,
      selectionAfterRanges: ranges.map((range) => ({ ...range })),
      label: `${rangesToLabel(ranges)} ${sheetIndexes.length} シート${homeClearLabel(normalizedMode)}`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangesToLabel(ranges)} の${homeClearLabel(normalizedMode)}を実行しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`クリアを実行できませんでした：${error.message || error}`);
  }
  return true;
}



function clearCellModelForHome(sheet, key, mode) {
  const existing = sheet.cells[key];
  if (!existing) return;
  if (mode === "all") {
    delete sheet.cells[key];
    return;
  }
  const cell = cloneCellModel(existing);
  if (mode === "contents") {
    cell.raw = null;
    cell.formula = null;
    cell.cached = null;
    cell.display = "";
    cell.dataTableFormula = "";
    delete cell.richText;
    delete cell.phonetics;
    delete cell.showPhonetic;
    delete cell.sparkline;
  } else if (mode === "formats") {
    cell.css = {};
    cell.hasFill = false;
    cell.hasFontColor = false;
    cell.borders = {};
    cell.numFmt = "";
    delete cell.richText;
  } else if (mode === "comments") {
    cell.comment = null;
    cell.note = null;
    cell.noteVisible = false;
  } else if (mode === "hyperlinks") {
    cell.hyperlink = null;
  }
  storeCellModelForSheet(sheet, key, cell);
}



function homeClearLabel(mode) {
  return {
    all: "すべてクリア",
    formats: "書式のクリア",
    contents: "数式と値のクリア",
    comments: "コメントとメモのクリア",
    hyperlinks: "ハイパーリンクのクリア",
  }[mode] || "クリア";
}



function handleFindSelectMenuAction(event) {
  const action = event.currentTarget.dataset.findSelectAction;
  closeHomeCommandMenu();
  if (action === "find") {
    openFindDialog();
  } else if (action === "replace") {
    openFindDialog({ mode: "replace" });
  } else if (action === "go-to") {
    openGoToDialog();
  } else if (action === "go-to-special") {
    openSpecialDialog();
  } else if (action === "select-formulas") {
    selectCellsByPredicate("数式", (cell) => Boolean(cell?.formula || cell?.dataTableFormula || (typeof cell?.raw === "string" && cell.raw.startsWith("="))));
  } else if (action === "select-comments") {
    selectCellsByPredicate("コメント", cellHasComment);
  } else if (action === "select-notes") {
    selectCellsByPredicate("メモ", cellHasNote);
  } else if (action === "select-conditional-formats") {
    selectCellsByPredicate("条件付き書式", cellHasConditionalFormat);
  } else if (action === "select-constants") {
    selectCellsByPredicate("定数", cellHasConstant);
  } else if (action === "select-validation") {
    selectCellsByPredicate("データの入力規則", (cell) => Boolean(cell?.validation));
  } else if (action === "select-objects") {
    enterObjectSelectionMode();
  }
}



function openGoToDialog(defaultReference = "") {
  closeGoToDialog();
  closeSpecialDialog();
  closeFindDialog();
  const visibleNames = (state.model?.definedNames || []).filter((definition) => !definition.hidden);
  const currentReference = defaultReference || (state.selected?.sheetIndex === state.activeSheetIndex ? `${columnName(state.selected.col)}${state.selected.row}` : "A1");
  const overlay = document.createElement("div");
  overlay.id = "goToDialogOverlay";
  overlay.className = "excel-dialog-overlay go-to-dialog-overlay";
  overlay.innerHTML = `
    <div id="goToDialog" class="excel-mini-dialog go-to-dialog" role="dialog" aria-modal="true" aria-labelledby="goToDialogTitle">
      <div class="excel-mini-titlebar" id="goToDialogTitle">ジャンプ</div>
      <div class="excel-mini-body go-to-dialog-body">
        <div class="go-to-list" role="listbox" aria-label="移動先">
          ${visibleNames.length
            ? visibleNames.map((definition) => `<button type="button" data-go-to-dialog-action="pick" data-go-to-reference="${escapeAttr(definition.name)}">${escapeHtml(definition.name)}</button>`).join("")
            : `<div class="go-to-empty">定義された名前はありません。</div>`}
        </div>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("参照先(R):")}</span><input type="text" data-go-to-field="reference" value="${escapeAttr(currentReference)}" /></label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-go-to-dialog-action="special">セル選択(S)...</button>
        <button class="excel-mini-button is-primary" type="button" data-go-to-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-go-to-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.append(overlay);
  overlay.querySelector("[data-go-to-field='reference']")?.focus({ preventScroll: true });
  overlay.querySelector("[data-go-to-field='reference']")?.select?.();
}



function handleGoToDialogAction(event) {
  const action = event.currentTarget.dataset.goToDialogAction;
  const dialog = event.currentTarget.closest("#goToDialog");
  if (action === "cancel") {
    closeGoToDialog();
  } else if (action === "special") {
    const reference = dialog?.querySelector("[data-go-to-field='reference']")?.value || "";
    closeGoToDialog();
    openSpecialDialog({ returnReference: reference });
  } else if (action === "pick") {
    const reference = event.currentTarget.dataset.goToReference || "";
    const input = dialog?.querySelector("[data-go-to-field='reference']");
    if (input) input.value = reference;
  } else if (action === "ok") {
    goToReferenceFromDialog(dialog);
  }
}



function handleGoToDialogKeydown(event) {
  if (isElementImeCompositionEvent(event)) return;
  if (event.key === "Escape") {
    event.preventDefault();
    closeGoToDialog();
  } else if (event.key === "Enter") {
    event.preventDefault();
    goToReferenceFromDialog(event.currentTarget);
  }
}



function goToReferenceFromDialog(dialog) {
  const reference = String(dialog?.querySelector("[data-go-to-field='reference']")?.value || "").trim();
  const target = resolveGoToReference(reference);
  if (!target) {
    showMessageBox("参照が正しくありません。", { title: "Cht WebSheet" });
    return;
  }
  closeGoToDialog();
  activateGoToTarget(target, `${target.sheet.name}!${rangeToLabel(target.range)} に移動しました。`);
}



function resolveGoToReference(reference) {
  const text = String(reference || "").trim().replace(/^=/, "");
  if (!text) return null;
  const active = activeSheet();
  const direct = validationQualifiedRange(text.replace(/\$/g, ""), active);
  if (direct) return goToTargetFromQualifiedRange(direct);
  const definition = validationDefinedNameFromFormulaText(text, active);
  if (definition?.refersTo) {
    const firstRef = firstDefinedNameReference(definition.refersTo);
    const definitionSheet = definition.scope === "sheet"
      ? state.model?.sheets?.[definition.sheetIndex] || active
      : active;
    const named = validationReferenceFromFormula(definitionSheet, firstRef);
    if (named) return goToTargetFromQualifiedRange(named);
  }
  const range = decodeRange(text.replace(/\$/g, ""));
  return range ? { sheet: active, sheetIndex: state.activeSheetIndex, range: { sheetIndex: state.activeSheetIndex, ...range } } : null;
}



function goToTargetFromQualifiedRange(qualified) {
  const sheetIndex = (state.model?.sheets || []).indexOf(qualified.sheet);
  if (sheetIndex < 0) return null;
  return { sheet: qualified.sheet, sheetIndex, range: { sheetIndex, ...qualified.range } };
}



function activateGoToTarget(target, statusMessage) {
  const sheet = target.sheet;
  state.activeSheetIndex = target.sheetIndex;
  ensureActiveSheetVisible();
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  selectRangesForHomeCommand([{ ...target.range, sheetIndex: target.sheetIndex }], statusMessage);
}



function closeGoToDialog() {
  document.querySelector("#goToDialogOverlay")?.remove();
}



function openSpecialDialog(options = {}) {
  closeSpecialDialog();
  closeGoToDialog();
  closeFindDialog();
  const overlay = document.createElement("div");
  overlay.id = "specialDialogOverlay";
  overlay.className = "excel-dialog-overlay special-dialog-overlay";
  overlay.innerHTML = `
    <div id="specialDialog" class="excel-mini-dialog special-dialog" role="dialog" aria-modal="true" aria-labelledby="specialDialogTitle" data-return-reference="${escapeAttr(options.returnReference || "")}">
      <div class="excel-mini-titlebar" id="specialDialogTitle">選択オプション</div>
      <div class="excel-mini-body special-dialog-body">
        ${specialOptionHtml("comments", "コメント(M)", "コメントが設定されているセル")}
        ${specialOptionHtml("constants", "定数(O)", "数式以外の値が入力されているセル")}
        ${specialOptionHtml("formulas", "数式(F)", "数式が入力されているセル")}
        ${specialOptionHtml("blanks", "空白セル(K)", "選択範囲内の空白セル")}
        ${specialOptionHtml("row-differences", "行内の相違(W)", "各行の先頭セルと値が異なるセル")}
        ${specialOptionHtml("column-differences", "列内の相違(C)", "各列の先頭セルと値が異なるセル")}
        ${specialOptionHtml("precedents", "参照元(P)", "選択セルの数式が参照しているセル")}
        ${specialOptionHtml("dependents", "参照先(D)", "選択セルを参照している数式セル")}
        ${specialOptionHtml("current-region", "アクティブ セル領域(R)", "アクティブ セルを含む連続した範囲")}
        ${specialOptionHtml("visible", "可視セル(Y)", "非表示の行と列を除いたセル")}
        ${specialOptionHtml("last-cell", "最後のセル(S)", "使用されている最後のセル")}
        ${specialOptionHtml("conditional-formats", "条件付き書式(T)", "条件付き書式が設定されているセル")}
        ${specialOptionHtml("validation", "データの入力規則(V)", "入力規則が設定されているセル")}
        ${specialOptionHtml("same-validation", "同じ入力規則(A)", "アクティブ セルと同じ入力規則のセル")}
        ${specialOptionHtml("objects", "オブジェクトの選択(B)", "図形、画像、グラフなどのオブジェクト")}
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-special-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-special-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.append(overlay);
  overlay.querySelector("input[name='specialOption']")?.focus({ preventScroll: true });
}



function specialOptionHtml(value, label, hint) {
  return `<label class="special-dialog-option"><input type="radio" name="specialOption" value="${escapeAttr(value)}" ${value === "comments" ? "checked" : ""} /><span>${accessKeyLabelHtml(label)}</span><small>${escapeHtml(hint)}</small></label>`;
}



function handleSpecialDialogAction(event) {
  const action = event.currentTarget.dataset.specialDialogAction;
  const dialog = event.currentTarget.closest("#specialDialog");
  if (action === "cancel") {
    closeSpecialDialog();
  } else if (action === "ok") {
    applySpecialSelection(dialog);
  }
}



function handleSpecialDialogKeydown(event) {
  if (isElementImeCompositionEvent(event)) return;
  if (event.key === "Escape") {
    event.preventDefault();
    closeSpecialDialog();
  } else if (event.key === "Enter") {
    event.preventDefault();
    applySpecialSelection(event.currentTarget);
  }
}



function applySpecialSelection(dialog) {
  const option = dialog?.querySelector("input[name='specialOption']:checked")?.value || "comments";
  closeSpecialDialog();
  if (option === "comments") selectCellsByPredicate("コメント", cellHasComment);
  else if (option === "constants") selectCellsByPredicate("定数", cellHasConstant);
  else if (option === "formulas") selectCellsByPredicate("数式", (cell) => Boolean(cell?.formula || cell?.dataTableFormula || (typeof cell?.raw === "string" && cell.raw.startsWith("="))));
  else if (option === "blanks") selectBlankCellsFromSelection();
  else if (option === "row-differences") selectDifferencesFromSelection("row");
  else if (option === "column-differences") selectDifferencesFromSelection("column");
  else if (option === "precedents") selectFormulaReferencesFromSelection("precedents");
  else if (option === "dependents") selectFormulaReferencesFromSelection("dependents");
  else if (option === "current-region") selectCurrentRegionFromHomeCommand();
  else if (option === "visible") selectVisibleCellsFromSelection();
  else if (option === "last-cell") selectLastUsedCellFromHomeCommand();
  else if (option === "conditional-formats") selectCellsByPredicate("条件付き書式", cellHasConditionalFormat);
  else if (option === "validation") selectCellsByPredicate("データの入力規則", (cell) => Boolean(cell?.validation));
  else if (option === "same-validation") selectSameValidationCellsFromSelection();
  else if (option === "objects") enterObjectSelectionMode();
}



function closeSpecialDialog() {
  document.querySelector("#specialDialogOverlay")?.remove();
}



function rangeForSpecialSelection() {
  const range = activeSelectionRange();
  if (!range) return null;
  if (!isSingleCellRange(range)) return range;
  const point = state.selected?.sheetIndex === state.activeSheetIndex
    ? state.selected
    : { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
  return currentRegionFromPoint(activeSheet(), point) || range;
}



function selectBlankCellsFromSelection() {
  const sheet = activeSheet();
  const range = rangeForSpecialSelection();
  if (!range) {
    showMessageBox("対象範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const ranges = [];
  forEachCellInRange(range, (row, col) => {
    if (!cellHasDataForRegion(sheet, row, col)) {
      ranges.push(normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col }));
    }
  });
  if (!ranges.length) {
    showMessageBox("該当する空白セルが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  selectRangesForHomeCommand(ranges, `空白セルを ${ranges.length} 個選択しました。`);
}



function selectDifferencesFromSelection(axis = "row") {
  const sheet = activeSheet();
  const range = rangeForSpecialSelection();
  if (!range) {
    showMessageBox("対象範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const byRow = axis === "row";
  if (byRow && range.left === range.right) {
    showMessageBox("行内の相違を調べるには、2 列以上の範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  if (!byRow && range.top === range.bottom) {
    showMessageBox("列内の相違を調べるには、2 行以上の範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const ranges = [];
  if (byRow) {
    for (let row = range.top; row <= range.bottom; row += 1) {
      const baseline = cellComparableValueForSpecial(sheet, row, range.left);
      for (let col = range.left + 1; col <= range.right; col += 1) {
        if (cellComparableValueForSpecial(sheet, row, col) !== baseline) {
          ranges.push(cellRangeForSpecial(row, col));
        }
      }
    }
  } else {
    for (let col = range.left; col <= range.right; col += 1) {
      const baseline = cellComparableValueForSpecial(sheet, range.top, col);
      for (let row = range.top + 1; row <= range.bottom; row += 1) {
        if (cellComparableValueForSpecial(sheet, row, col) !== baseline) {
          ranges.push(cellRangeForSpecial(row, col));
        }
      }
    }
  }
  if (!ranges.length) {
    showMessageBox("該当する相違セルが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const label = byRow ? "行内の相違" : "列内の相違";
  selectRangesForHomeCommand(uniqueSelectionRanges(ranges), `${label}セルを ${ranges.length} 個選択しました。`);
}



function cellComparableValueForSpecial(sheet, row, col) {
  const sheetId = getSheetId(sheet?.name);
  const cell = sheet?.cells?.[cellKey(row, col)];
  if (!cell) return "";
  return String(getDisplayForCell(sheet, sheetId, row, col) ?? "");
}



function cellRangeForSpecial(row, col) {
  return normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col });
}



function selectCurrentRegionFromHomeCommand() {
  const range = rangeForSpecialSelection();
  if (!range) {
    showMessageBox("アクティブ セル領域が見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  selectSingleRange(range, `${rangeToLabel(range)} を選択しました。`);
}



function selectFormulaReferencesFromSelection(mode = "precedents") {
  const sheet = activeSheet();
  const ranges = activeSelectionRanges();
  if (!sheet || !ranges.length) {
    showMessageBox("対象範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const selectedRanges = ranges.map((range) => ({ ...range, sheetIndex: state.activeSheetIndex }));
  const resultRanges = mode === "dependents"
    ? dependentFormulaRangesForSelection(sheet, selectedRanges)
    : precedentRangesForSelection(sheet, selectedRanges);
  if (!resultRanges.length) {
    const target = mode === "dependents" ? "参照先" : "参照元";
    showMessageBox(`該当する${target}セルが見つかりません。`, { title: "Cht WebSheet" });
    return;
  }
  const target = mode === "dependents" ? "参照先" : "参照元";
  selectRangesForHomeCommand(resultRanges, `${target}セルを ${resultRanges.length} 個選択しました。`);
}



function precedentRangesForSelection(sheet, selectedRanges) {
  const ranges = [];
  selectedRanges.forEach((range) => {
    forEachCellInRange(range, (row, col) => {
      ranges.push(...formulaReferencesForCell(sheet, row, col).filter((reference) => reference.sheetIndex === state.activeSheetIndex).map((reference) => reference.range));
    });
  });
  return uniqueSelectionRanges(ranges);
}



function dependentFormulaRangesForSelection(sheet, selectedRanges) {
  return uniqueSelectionRanges(dependentFormulaTargetsForSelection(sheet, selectedRanges).map((target) => target.range));
}



function dependentFormulaTargetsForSelection(sheet, selectedRanges) {
  const targets = [];
  const seen = new Set();
  Object.entries(sheet?.cells || {}).forEach(([key, cell]) => {
    if (!cellHasFormula(cell)) return;
    const [row, col] = parseCellKey(key);
    const references = formulaReferencesForCell(sheet, row, col).filter((reference) => reference.sheetIndex === state.activeSheetIndex);
    if (!references.some((reference) => selectedRanges.some((selectedRange) => rangesIntersect(reference.range, selectedRange)))) return;
    const range = cellRangeForSpecial(row, col);
    const uniqueKey = `${row}:${col}:${rangeToLabel(range)}`;
    if (seen.has(uniqueKey)) return;
    seen.add(uniqueKey);
    targets.push({ row, col, range, references });
  });
  return targets;
}



function formulaReferencesForCell(sheet, row, col) {
  const cell = sheet?.cells?.[cellKey(row, col)];
  const raw = getCellRawInput(sheet, row, col);
  const text = String(raw || "").trim();
  if (!cellHasFormula(cell) && !text.startsWith("=")) return [];
  return parseFormulaReferences(text.startsWith("=") ? text : `=${text}`, sheet);
}



function parseFormulaReferences(formula, baseSheet) {
  const source = String(formula || "").replace(/^=/, "");
  const references = [];
  for (const token of formulaReferenceTokensInText(source)) {
    const sheetToken = token.sheetToken || "";
    const targetSheet = sheetToken
      ? (state.model?.sheets || []).find((item) => item.name === unquoteSheetName(sheetToken)) || null
      : baseSheet;
    if (!targetSheet) continue;
    const sheetIndex = (state.model?.sheets || []).indexOf(targetSheet);
    if (sheetIndex < 0) continue;
    const rangeText = token.body.replace(/\$/g, "");
    const range = decodeExcelRangeReference(rangeText, targetSheet);
    if (range) references.push({ sheet: targetSheet, sheetIndex, range: { sheetIndex, ...range } });
  }
  return references;
}



function selectVisibleCellsFromSelection() {
  const sheet = activeSheet();
  const range = rangeForSpecialSelection();
  if (!range) {
    showMessageBox("対象範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const ranges = [];
  forEachCellInRange(range, (row, col) => {
    if (sheet.rows[row - 1]?.hidden || sheet.columns[col - 1]?.hidden) return;
    ranges.push(normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col }));
  });
  if (!ranges.length) {
    showMessageBox("該当する可視セルが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  selectRangesForHomeCommand(ranges, `可視セルを ${ranges.length} 個選択しました。`);
}



function selectLastUsedCellFromHomeCommand() {
  const used = sheetUsedRange(activeSheet());
  if (!used) {
    showMessageBox("最後のセルが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const range = normalizeRange({ sheetIndex: state.activeSheetIndex, row: used.bottom, col: used.right }, { sheetIndex: state.activeSheetIndex, row: used.bottom, col: used.right });
  selectSingleRange(range, `${rangeToLabel(range)} を選択しました。`);
}



function selectSameValidationCellsFromSelection() {
  const sheet = activeSheet();
  const selected = state.selected?.sheetIndex === state.activeSheetIndex
    ? state.selected
    : { sheetIndex: state.activeSheetIndex, row: activeSelectionRange()?.top || 1, col: activeSelectionRange()?.left || 1 };
  const activeRule = sheet?.cells?.[cellKey(selected.row, selected.col)]?.validation;
  if (!activeRule) {
    showMessageBox("アクティブ セルにはデータの入力規則が設定されていません。", { title: "Cht WebSheet" });
    return;
  }
  const key = validationComparisonKey(activeRule);
  const ranges = matchingCellRanges(sheet, (cell) => Boolean(cell?.validation) && validationComparisonKey(cell.validation) === key);
  if (!ranges.length) {
    showMessageBox("同じ入力規則のセルが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  selectRangesForHomeCommand(ranges, `同じ入力規則のセルを ${ranges.length} 個選択しました。`);
}



function validationComparisonKey(rule) {
  const clone = cloneJsonLikeValue(rule) || {};
  delete clone.anchor;
  return stableObjectKey(clone);
}



function uniqueSelectionRanges(ranges) {
  const seen = new Set();
  const unique = [];
  (ranges || []).forEach((range) => {
    if (!range || range.sheetIndex !== state.activeSheetIndex) return;
    const normalized = {
      sheetIndex: state.activeSheetIndex,
      top: Math.min(range.top, range.bottom),
      left: Math.min(range.left, range.right),
      bottom: Math.max(range.top, range.bottom),
      right: Math.max(range.left, range.right),
      ...(range.fullRow ? { fullRow: true } : {}),
      ...(range.fullColumn ? { fullColumn: true } : {}),
    };
    const expanded = normalized.fullRow || normalized.fullColumn ? normalized : expandRangeForMerges(activeSheet(), normalized) || normalized;
    const key = rangeToLabel(expanded);
    if (seen.has(key)) return;
    seen.add(key);
    unique.push(expanded);
  });
  unique.sort((left, right) => left.top - right.top || left.left - right.left || left.bottom - right.bottom || left.right - right.right);
  return unique;
}



function enterObjectSelectionMode() {
  if (!state.model) return;
  closeHomeCommandMenu();
  closeCellContextMenu();
  closeRibbonColorPalette();
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeShapeFormatMenu();
  closeAutoFilterMenu();
  $sheetHost.find(".sheet-cell[contenteditable='true']")[0]?.blur?.();
  clearPictureCropState();
  state.objectSelectionMode = true;
  removeObjectSelectionMarquee();
  document.body.classList.add("is-object-selection-mode");
  setStatus("オブジェクトの選択モードです。選択する範囲をドラッグしてください。Escで終了します。");
}



function exitObjectSelectionMode(statusMessage = "") {
  state.objectSelectionMode = false;
  removeObjectSelectionMarquee();
  document.body.classList.remove("is-object-selection-mode");
  if (statusMessage) setStatus(statusMessage);
}



function startObjectSelectionModeDrag(event) {
  if (!state.objectSelectionMode || event.button !== 0) return;
  if (event.target?.closest?.(".excel-dialog-overlay, .format-dialog-overlay, .cell-structure-dialog-overlay, .home-command-menu, .cell-context-menu, .ribbon-shape-menu, .sheet-footer")) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  commitActiveCellEditForExternalPointer(event);
  const marquee = document.createElement("div");
  marquee.className = "object-selection-marquee";
  document.body.append(marquee);
  state.objectSelectionDrag = {
    startX: event.clientX,
    startY: event.clientY,
    currentX: event.clientX,
    currentY: event.clientY,
    marquee,
  };
  updateObjectSelectionMarquee();
}



function updateObjectSelectionModeDrag(event) {
  if (!state.objectSelectionDrag) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  state.objectSelectionDrag.currentX = event.clientX;
  state.objectSelectionDrag.currentY = event.clientY;
  updateObjectSelectionMarquee();
}



function finishObjectSelectionModeDrag(event) {
  const drag = state.objectSelectionDrag;
  if (!drag) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  const rect = normalizedClientRectFromObjectSelectionDrag(drag);
  const ids = objectIdsIntersectingClientRect(rect);
  removeObjectSelectionMarquee();
  state.objectSelectionMode = false;
  document.body.classList.remove("is-object-selection-mode");
  if (!ids.length) {
    setSelectedImageIds([], null);
    renderSheet();
    updateSelectionUi();
    setStatus("範囲内にオブジェクトが見つかりません。");
    return;
  }
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionStart = null;
  state.selectionAnchor = null;
  state.selectionDrag = null;
  setSelectedImageIds(ids, ids[ids.length - 1]);
  $formulaInput.prop("disabled", true).val("");
  renderSheet();
  updateSelectionUi();
  setStatus(`${ids.length} 個のオブジェクトを選択しました。`);
}



function cancelObjectSelectionModeDrag() {
  if (!state.objectSelectionMode && !state.objectSelectionDrag) return;
  exitObjectSelectionMode("オブジェクトの選択モードを終了しました。");
}



function removeObjectSelectionMarquee() {
  state.objectSelectionDrag?.marquee?.remove?.();
  state.objectSelectionDrag = null;
}



function normalizedClientRectFromObjectSelectionDrag(drag) {
  const left = Math.min(drag.startX, drag.currentX);
  const top = Math.min(drag.startY, drag.currentY);
  const right = Math.max(drag.startX, drag.currentX);
  const bottom = Math.max(drag.startY, drag.currentY);
  return { left, top, right, bottom, width: right - left, height: bottom - top };
}



function updateObjectSelectionMarquee() {
  const drag = state.objectSelectionDrag;
  if (!drag?.marquee) return;
  const rect = normalizedClientRectFromObjectSelectionDrag(drag);
  drag.marquee.style.left = `${roundCssNumber(rect.left)}px`;
  drag.marquee.style.top = `${roundCssNumber(rect.top)}px`;
  drag.marquee.style.width = `${roundCssNumber(Math.max(1, rect.width))}px`;
  drag.marquee.style.height = `${roundCssNumber(Math.max(1, rect.height))}px`;
}



function objectIdsIntersectingClientRect(rect) {
  const sheet = activeSheet();
  if (!sheet?.images?.length || !$sheetHost[0]) return [];
  const candidates = sheet.images
    .filter((image) => image && !image.hidden)
    .map((image) => {
      const escapedId = cssEscape(image.id);
      const element = $sheetHost[0].querySelector(`.sheet-image[data-image-id="${escapedId}"], .sheet-image-link[data-image-id="${escapedId}"]`);
      return element ? { image, rect: element.getBoundingClientRect() } : null;
    })
    .filter((entry) => entry && clientRectsIntersect(rect, entry.rect))
    .sort((left, right) => worksheetObjectLayerZIndex(sheet, left.image) - worksheetObjectLayerZIndex(sheet, right.image));
  return candidates.map((entry) => entry.image.id);
}



function clientRectsIntersect(left, right) {
  return left.left <= right.right && left.right >= right.left && left.top <= right.bottom && left.bottom >= right.top;
}



function selectObjectsFromHomeCommand() {
  const sheet = activeSheet();
  const ids = (sheet.images || []).filter((image) => !image.hidden).map((image) => image.id);
  if (!ids.length) {
    showMessageBox("該当するオブジェクトが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionStart = null;
  state.selectionAnchor = null;
  state.selectionDrag = null;
  setSelectedImageIds(ids, ids[0]);
  $formulaInput.prop("disabled", true).val("");
  renderSheet();
  updateSelectionUi();
  setStatus(`${ids.length} 個のオブジェクトを選択しました。`);
}



function openFindDialog(options = {}) {
  closeFindDialog();
  closeGoToDialog();
  closeSpecialDialog();
  closeHomeCommandMenu();
  const mode = options.mode === "replace" ? "replace" : "find";
  const scope = normalizeFindScope(options.scope || state.lastFind?.options?.scope);
  state.findDialog = { query: "", mode, scope };
  const overlay = document.createElement("div");
  overlay.id = "findDialogOverlay";
  overlay.className = "excel-dialog-overlay find-dialog-overlay";
  overlay.innerHTML = `
    <div id="findDialog" class="excel-mini-dialog find-dialog" role="dialog" aria-modal="true" aria-labelledby="findDialogTitle">
      <div class="excel-mini-titlebar" id="findDialogTitle">検索と置換</div>
      <div class="excel-mini-body find-dialog-body">
        <div class="find-dialog-tabs" role="tablist" aria-label="検索と置換">
          <button class="find-dialog-tab${mode === "find" ? " is-active" : ""}" type="button" role="tab" aria-selected="${mode === "find"}" data-find-dialog-action="show-find">検索</button>
          <button class="find-dialog-tab${mode === "replace" ? " is-active" : ""}" type="button" role="tab" aria-selected="${mode === "replace"}" data-find-dialog-action="show-replace">置換</button>
        </div>
        <label class="excel-dialog-field find-dialog-query"><span class="excel-dialog-caption">${accessKeyLabelHtml("検索する文字列(N):")}</span><input type="text" data-find-field="query" /></label>
        <label class="excel-dialog-field find-dialog-replace${mode === "replace" ? "" : " hidden"}"><span class="excel-dialog-caption">${accessKeyLabelHtml("置換後の文字列(E):")}</span><input type="text" data-find-field="replacement" /></label>
        <label class="excel-dialog-field find-dialog-scope"><span class="excel-dialog-caption">${accessKeyLabelHtml("検索場所(H):")}</span><select data-find-field="scope"><option value="sheet"${scope === "sheet" ? " selected" : ""}>シート</option><option value="workbook"${scope === "workbook" ? " selected" : ""}>ブック</option></select></label>
        <div class="find-dialog-options">
          <label><input type="checkbox" data-find-field="caseSensitive" />大文字と小文字を区別する(C)</label>
          <label><input type="checkbox" data-find-field="wholeCell" />セル内容が完全に同一であるものを検索する(O)</label>
        </div>
        <div class="find-dialog-results hidden" data-find-results>
          <div class="find-dialog-results-summary" data-find-results-summary></div>
          <div class="find-dialog-results-head" aria-hidden="true"><span>シート</span><span>セル</span><span>値</span></div>
          <div class="find-dialog-results-list" role="listbox" aria-label="検索結果" data-find-results-list></div>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button find-replace-action${mode === "replace" ? "" : " hidden"}" type="button" data-find-dialog-action="replace-all">すべて置換(A)</button>
        <button class="excel-mini-button find-replace-action${mode === "replace" ? "" : " hidden"}" type="button" data-find-dialog-action="replace-one">置換(R)</button>
        <button class="excel-mini-button" type="button" data-find-dialog-action="find-all">すべて検索(I)</button>
        <button class="excel-mini-button is-primary" type="button" data-find-dialog-action="find-next">次を検索(F)</button>
        <button class="excel-mini-button" type="button" data-find-dialog-action="close">閉じる</button>
      </div>
    </div>`;
  document.body.append(overlay);
  overlay.querySelector("[data-find-field='query']")?.focus({ preventScroll: true });
}



function handleFindDialogAction(event) {
  const action = event.currentTarget.dataset.findDialogAction;
  if (action === "close" || action === "cancel") {
    closeFindDialog();
    return;
  }
  if (action === "find-next") findNextFromDialog(event.currentTarget.closest("#findDialog"));
  if (action === "find-all") findAllFromDialog(event.currentTarget.closest("#findDialog"));
  if (action === "select-result") selectFindResultFromDialog(event.currentTarget.closest("#findDialog"), event.currentTarget);
  if (action === "replace-one") replaceFromDialog(event.currentTarget.closest("#findDialog"), { all: false });
  if (action === "replace-all") replaceFromDialog(event.currentTarget.closest("#findDialog"), { all: true });
  if (action === "show-find") setFindDialogMode(event.currentTarget.closest("#findDialog"), "find");
  if (action === "show-replace") setFindDialogMode(event.currentTarget.closest("#findDialog"), "replace");
}



function setFindDialogMode(dialog, mode = "find") {
  if (!dialog) return;
  const replaceMode = mode === "replace";
  state.findDialog = { ...(state.findDialog || {}), mode: replaceMode ? "replace" : "find" };
  dialog.querySelectorAll(".find-dialog-tab").forEach((tab) => {
    const isActive = tab.dataset.findDialogAction === (replaceMode ? "show-replace" : "show-find");
    tab.classList.toggle("is-active", isActive);
    tab.setAttribute("aria-selected", String(isActive));
  });
  dialog.querySelector(".find-dialog-replace")?.classList.toggle("hidden", !replaceMode);
  dialog.querySelectorAll(".find-replace-action").forEach((button) => button.classList.toggle("hidden", !replaceMode));
  dialog.querySelector("[data-find-field='query']")?.focus({ preventScroll: true });
}

export {
  activateCommandMenuItem,
  activateGoToTarget,
  activeCommandMenuForGlobalAccessKey,
  applyHomeJustifyFill,
  applyHomeSeriesDialog,
  applyHomeWorksheetFillDialog,
  applyPageSetupDialog,
  applySpecialSelection,
  autoSumCellLooksSummable,
  autoSumFormulaForCell,
  autoSumMenuButtonHtml,
  autoSumOperationsForRanges,
  beginInkDrawingMode,
  beginInkEraserMode,
  cancelInkDrawingMode,
  cancelObjectSelectionModeDrag,
  cellComparableValueForSpecial,
  cellRangeForSpecial,
  clearCellModelForHome,
  clearInkObjects,
  clearMenuButtonHtml,
  clearSelectionFromHome,
  clearSelectionFromHomeSheetGroup,
  clientRectsIntersect,
  closeActiveCommandMenu,
  closeGoToDialog,
  closeHomeCommandMenu,
  closeHomeCommandMenuOnEscape,
  closeHomeCommandMenuOnOutsidePointer,
  closeHomeSeriesDialog,
  closeHomeWorksheetFillDialog,
  closeSpecialDialog,
  collectHomeSeriesDialogValues,
  columnWidthForIndex,
  commandMenuEscapeBlockedByDialog,
  commandMenuItemCenter,
  commandMenuItems,
  consumeCommandMenuKey,
  contiguousAutoSumSource,
  convertSelectedInkToFreeformShape,
  copyWorksheetFillCell,
  createInkDrawingPreview,
  createInkStrokeFromPoints,
  customPathFromSheetPoints,
  dependentFormulaRangesForSelection,
  dependentFormulaTargetsForSelection,
  distancePointToSegment,
  enhanceCommandMenuAccessKeys,
  enterObjectSelectionMode,
  eraseInkAtClientPoint,
  eraseInkPointAtSheetPoint,
  eraseInkStrokeAtSheetPoint,
  eraseInkStrokePartAtPoint,
  exitObjectSelectionMode,
  fillMenuButtonHtml,
  fillSelectionFromHome,
  fillSelectionFromHomeSheetGroup,
  findSelectMenuButtonHtml,
  finishInkDrawing,
  finishInkEraser,
  finishObjectSelectionModeDrag,
  focusCommandMenuItem,
  formulaReferencesForCell,
  goToReferenceFromDialog,
  goToTargetFromQualifiedRange,
  handleFindDialogAction,
  handleFindSelectMenuAction,
  handleFormulaCalculationMenuAction,
  handleGlobalCommandMenuAccessKey,
  handleGoToDialogAction,
  handleGoToDialogKeydown,
  handleHomeAutoSumMenuAction,
  handleHomeClearMenuAction,
  handleHomeFillMenuAction,
  handleHomeSeriesDialogAction,
  handleHomeSeriesDialogKeydown,
  handleHomeSeriesDialogSync,
  handleHomeSortFilterMenuAction,
  handleHomeWorksheetFillDialogAction,
  handleHomeWorksheetFillDialogKeydown,
  handlePageSetupTabClick,
  handleRibbonCommandMenuKeydown,
  handleSpecialDialogAction,
  handleSpecialDialogKeydown,
  hitTestInkStroke,
  homeClearLabel,
  homeCommandActionButtonHtml,
  homeCommandMenuButtonHtml,
  homeFillPlanForRange,
  homeFillRanges,
  homeFillSourcePoint,
  homeJustifyRangePixelWidth,
  homeJustifyTargets,
  homeJustifyTextChunks,
  homeSeriesDateValue,
  homeSeriesSeedValue,
  homeSeriesValueForOffset,
  inkStrokeBoundingBox,
  inkStrokeColor,
  inkStrokeDistanceToPoint,
  insertInkPathFromPoints,
  interpolateInkPoint,
  mapPrintAxisCoordinate,
  nextCommandMenuItemIndexByArrow,
  normalizedClientRectFromObjectSelectionDrag,
  objectIdsIntersectingClientRect,
  openFindDialog,
  openFindSelectMenu,
  openGoToDialog,
  openHomeAutoSumMenu,
  openHomeClearMenu,
  openHomeCommandMenu,
  openHomeFillMenu,
  openHomeSeriesDialog,
  openHomeSortFilterMenu,
  openHomeWorksheetFillDialog,
  openSpecialDialog,
  pageSetupReferenceBangIndex,
  pageSetupReferenceParts,
  parseFormulaReferences,
  parsePageSetupRangeLabel,
  pixelPointInsideImage,
  pointListBoundingBox,
  precedentRangesForSelection,
  printAxisSegments,
  printColumnAxisSegments,
  printHeadingOffsets,
  printImageHtml,
  printObjectPositionAttrs,
  printRangePixelBounds,
  printRowAxisSegments,
  printSheetImagesHtml,
  printSheetInkHtml,
  printSheetObjectsHtml,
  rangeBottomRowIsBlank,
  rangeForSpecialSelection,
  rangeHasSummableCell,
  rangeRightColumnIsBlank,
  removeObjectSelectionMarquee,
  resolveGoToReference,
  segmentCircleIntersections,
  selectBlankCellsFromSelection,
  selectCurrentRegionFromHomeCommand,
  selectDifferencesFromSelection,
  selectFormulaReferencesFromSelection,
  selectInkObjectsInsideLasso,
  selectLastUsedCellFromHomeCommand,
  selectObjectsFromHomeCommand,
  selectSameValidationCellsFromSelection,
  selectVisibleCellsFromSelection,
  setFindDialogMode,
  setInkColor,
  setInkWidth,
  simplifyInkPoints,
  specialOptionHtml,
  startInkDrawing,
  startObjectSelectionModeDrag,
  svgPathDFromSheetPoints,
  syncDrawRibbonButtons,
  syncHomeSeriesDialog,
  syncPrintObjectLayerPositions,
  uniqueSelectionRanges,
  updateInkDrawing,
  updateInkDrawingPreview,
  updateObjectSelectionMarquee,
  updateObjectSelectionModeDrag,
  validationComparisonKey,
  writeHomeFillCells,
  writeHomeSeriesCell,
  writeHomeSeriesLine,
};
