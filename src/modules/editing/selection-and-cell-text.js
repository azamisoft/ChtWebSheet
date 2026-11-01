// Feature module; shared bindings are imported explicitly.
import {
  $,
  $formulaExpandButton,
  $formulaInput,
  $ribbonPanels,
  $selectedCellLabel,
  $sheetHost,
  CELL_EDIT_MIN_PADDING_TOP,
  CELL_EDIT_PADDING_BOTTOM,
  CELL_EDIT_TRAILING_LINE_SENTINEL,
  COLUMN_HEADER_HEIGHT,
  CUSTOM_SHAPE_TYPE,
  DEFAULT_SHAPE_TYPE,
  EMU_PER_PIXEL,
  IMAGE_RENDER_BUFFER_PX,
  VIRTUAL_ROW_OVERSCAN,
  WORKSHEET_OBJECT_LAYER_BASE_Z,
  WORKSHEET_OBJECT_LAYER_STEP_Z,
  WORKSHEET_OBJECT_SELECTION_Z_OFFSET,
  state,
} from "../../app-core.js";
import {
  workbookQueries,
} from "../data/connections-and-queries.js";
import {
  flashFillRangeHasSourceData,
  flashFillTargetRangeForSelection,
  isAutoFilterHeaderCell,
  normalizeAutoFilterCriterion,
} from "../data/filter-and-text-tools.js";
import {
  forecastNumericValue,
  forecastTimelineValue,
  outlineAxisForRange,
  outlineAxisSheetAllows,
  outlineLevelAt,
} from "../data/forecast-and-statistics.js";
import {
  selectRangesForHomeCommand,
} from "../data/query-editor.js";
import {
  loadRecentDataSources,
} from "../data/query-transforms.js";
import {
  autoFilterSortKeyForHeaderCell,
  autoFilterStateForHeaderCell,
  autoFilterStateHasCriteria,
  currentRegionFromPoint,
  sheetUsedRange,
  sortStateHasKeys,
  visibleTableAutoFilters,
  visibleTableSortStates,
} from "../data/sort-and-autofilter.js";
import {
  canCreateWorkbookRelationship,
  hasValidationInvalidMark,
} from "../data/validation-and-consolidation.js";
import {
  chartSvg,
  outlineDetailTargetFromRange,
} from "../drawing/charts-and-insertion.js";
import {
  calloutBodySelectionBounds,
  calloutGeometry,
  clampNumber,
  cssColorForSvg,
  isCalloutShapeType,
  isLineShapeObject,
  normalizeRotationDegrees,
  normalizeShapeAdjustments,
  normalizeShapeCustomPath,
  normalizeShapeGlow,
  normalizeShapeLineCompound,
  normalizeShapeLineDash,
  normalizeShapeLinePoints,
  normalizeShapeOutlineWidth,
  normalizeShapeReflection,
  normalizeShapeSoftEdges,
  normalizeShapeTextMargin,
  normalizeShapeTextOutlineWidth,
  normalizeShapeTextRuns,
  roundRectCornerScale,
  shapeDescriptorForType,
  shapePlainTextFromRuns,
} from "../drawing/shape-rendering.js";
import {
  rotatePointAround,
  unionPixelBoxes,
} from "../drawing/transforms-and-connectors.js";
import {
  clearAutoFillOptions,
  isEnterKeyEvent,
  normalizeCellEditInlineCss,
} from "./autofill.js";
import {
  normalizeColorValue,
  setRibbonUnderlinePreset,
  svgDataUrl,
  underlinePresetForTextDecoration,
  underlinePresetForValue,
} from "./ink-and-clipboard.js";
import {
  isElementImeCompositionEvent,
} from "./keyboard-and-paste.js";
import {
  activeFormulaReferenceEntriesForActiveSheet,
  activeFormulaReferenceEntryForCell,
  activeSelectionRange,
  activeSelectionRanges,
  cellElementForPoint,
  cellPointFromElement,
  cellPointFromPointerEvent,
  findMergeForCell,
  formatPainterPointFromPointerEvent,
  getCellEditInput,
  getCellRawInput,
  getDisplayForCell,
  isActiveCell,
  isCellInActiveSelection,
  isCellInRange,
  isSingleCellRange,
  materializeSelectionRange,
  rangeToLabel,
  selectableCellPoint,
  selectionRangeFromPoints,
  sheetIndexFromPixel,
  sheetPixelFromClientPosition,
  updateSelectionAutoScroll,
  updateSelectionUi,
} from "./pointers-and-formula-evaluation.js";
import {
  cellEffectiveCss,
  cloneDefinedNames,
  cloneJsonLikeValue,
  normalizeInkStrokes,
  pushHistory,
  sheetDisplayCell,
} from "./structure-and-history.js";
import {
  formulaReferenceColorClass,
} from "../formula/autocomplete.js";
import {
  clearFormulaReferenceState,
  closeFormulaAutocomplete,
  extendFormulaHeaderReferenceSelection,
  startFormulaHeaderReferenceSelection,
  stopFormulaReferenceAutoScroll,
  updateFormulaReferenceUi,
} from "../formula/dialogs-and-names.js";
import {
  applyFormatPainterAtPoint,
  cellShowsValidationDropdown,
  frozenCellClassNames,
  frozenCellStickyCss,
  isFrozenPaneCell,
  rangeIntersects,
  rangeIntersectsRenderBands,
  rangesIntersect,
  renderSheet,
  repaintActiveSheetCanvas,
  richTextRunsForDisplay,
} from "../rendering/grid-and-canvas.js";
import {
  cellHasPhonetics,
  hyperlinkFromCellClickEvent,
  normalizeCommentText,
  numberCategoryFromFormat,
  openHyperlinkTarget,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  columnName,
  escapeAttr,
  escapeHtml,
  setStatus,
  showMessageBox,
  styleObjectToString,
} from "../ui/help-and-utilities.js";
import {
  inkStrokeBoundingBox,
  resolveGoToReference,
  svgPathDFromSheetPoints,
} from "../ui/printing-and-home-tools.js";
import {
  absoluteRangeReference,
  activeCellEditElementForFormatting,
  cancelIdleViewportTrim,
  cancelViewportRender,
  definedNameDuplicateMessage,
  definedNameHasConflict,
  ensureSheetSize,
  parseCellKey,
  rowHeaderWidthForSheet,
  upsertDefinedName,
  validateDefinedName,
} from "../ui/ribbon-and-links.js";
import {
  ensureEditableRangeForProtection,
  ensureSelectableRangeForProtection,
  isSheetProtected,
  rangeHasLockedCellsForProtection,
  sheetProtectionAllows,
} from "../ui/sheet-protection-and-translation.js";
import {
  formulaAuditClassForCell,
  sheetHeadingsVisible,
} from "../ui/views-and-formula-tools.js";
import {
  buildFormulaEngine,
  ensureActiveSheetVisible,
  fontFamilyForExcel,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  normalizePictureCrop,
} from "../workbook/ooxml-drawings.js";
import {
  activeSheet,
  cellDisplayText,
  forceSyncHorizontalScrollbarFromHost,
  getCellFormula,
  getSheetId,
  textOverflowWidth,
  visibleGridRange,
  widthBetweenColumns,
} from "../workbook/saving-and-export.js";
import {
  cachedWorksheetImageRenderSource,
  clampOutlineLevel,
  normalizeNote,
  normalizePhoneticRuns,
  rangeIntersectsWorksheetTable,
} from "../workbook/tables-and-validation.js";
import {
  rangeCellCount,
} from "../workbook/views-and-macros.js";


function phoneticRunsForLine(cell, line) {
  if (!cell?.showPhonetic || !cell?.phonetics?.length || !line) return [];
  return normalizePhoneticRuns(cell.phonetics, line);
}



function paintCanvasPhonetics(context, css, line, x, align, y, runs, fontSize) {
  const textLeft = canvasAlignedTextLeft(context, line, x, align);
  const placements = runs
    .map((run) => {
      const start = Math.max(0, Math.min(line.length, run.start));
      const end = Math.max(start, Math.min(line.length, run.end));
      const base = line.slice(start, end);
      if (!base) return null;
      const baseLeft = textLeft + context.measureText(line.slice(0, start)).width;
      const baseWidth = context.measureText(base).width;
      return { text: run.text, center: baseLeft + baseWidth / 2 };
    })
    .filter(Boolean);
  if (!placements.length) return;

  const previousFont = context.font;
  const previousAlign = context.textAlign;
  const previousBaseline = context.textBaseline;
  context.font = canvasFont(css, fontSize);
  context.textAlign = "left";
  context.textBaseline = "top";
  placements.forEach((placement) => {
    const width = context.measureText(placement.text).width;
    context.fillText(placement.text, placement.center - width / 2, y);
  });
  context.font = previousFont;
  context.textAlign = previousAlign;
  context.textBaseline = previousBaseline;
}



function canvasAlignedTextLeft(context, line, x, align) {
  const width = context.measureText(line).width;
  if (align === "right" || align === "end") return x - width;
  if (align === "center") return x - width / 2;
  return x;
}



function canvasFontDisplayScale(sheet) {
  const scale = Number(sheet?.displayScale || 1);
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}



function canvasFontSize(fontSize, scale = 1) {
  const value = Number.parseFloat(fontSize);
  const base = !Number.isFinite(value) || value <= 0 ? 13 : String(fontSize || "").trim().endsWith("pt") ? value * 96 / 72 : value;
  const safeScale = Number.isFinite(Number(scale)) && Number(scale) > 0 ? Number(scale) : 1;
  return Math.max(1, base * safeScale);
}



function canvasFont(css, fontSize = canvasFontSize(css?.fontSize)) {
  const style = css?.fontStyle === "italic" ? "italic " : "";
  const weight = css?.fontWeight ? `${css.fontWeight} ` : "";
  const family = css?.fontFamily || 'Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
  return `${style}${weight}${fontSize}px ${family}`;
}



function canvasTextLineHeight(fontSize) {
  return Math.max(10, fontSize * 1.2);
}



function canvasLineBaseline(context, line, lineTop, lineHeight, fontSize) {
  const metrics = context.measureText(line || "Mg");
  const ascent = Number(metrics.actualBoundingBoxAscent);
  const descent = Number(metrics.actualBoundingBoxDescent);
  const fallbackAscent = fontSize * 0.78;
  const fallbackDescent = fontSize * 0.22;
  const safeAscent = Number.isFinite(ascent) && ascent > 0 ? ascent : fallbackAscent;
  const safeDescent = Number.isFinite(descent) && descent >= 0 ? descent : fallbackDescent;
  const inkHeight = Math.max(1, safeAscent + safeDescent);
  return lineTop + Math.max(0, (lineHeight - inkHeight) / 2) + safeAscent;
}



function wrapCanvasText(context, text, maxWidth, font) {
  context.save();
  context.font = font;
  const hardLines = String(text).split(/\r\n|\r|\n/);
  const lines = [];
  hardLines.forEach((hardLine) => {
    if (!hardLine) {
      lines.push("");
      return;
    }
    let current = "";
    Array.from(hardLine).forEach((char) => {
      const candidate = current + char;
      if (current && context.measureText(candidate).width > maxWidth) {
        lines.push(current);
        current = char;
      } else {
        current = candidate;
      }
    });
    lines.push(current);
  });
  context.restore();
  return lines.slice(0, 200);
}



function paintCanvasUnderline(context, line, x, y, lineHeight, doubleUnderline = false) {
  const width = context.measureText(line).width;
  let start = x;
  if (context.textAlign === "center") start = x - width / 2;
  if (context.textAlign === "right") start = x - width;
  const underlineY = y + Math.max(1, Math.min(3, lineHeight * 0.12));
  const underlineGap = Math.max(2, Math.min(3, lineHeight * 0.16));
  context.save();
  context.strokeStyle = context.fillStyle;
  context.lineWidth = 1;
  context.beginPath();
  context.moveTo(start, underlineY);
  context.lineTo(start + width, underlineY);
  if (doubleUnderline) {
    context.moveTo(start, underlineY + underlineGap);
    context.lineTo(start + width, underlineY + underlineGap);
  }
  context.stroke();
  context.restore();
}



function paintCanvasStrikethrough(context, line, x, y, lineHeight) {
  const width = context.measureText(line).width;
  let start = x;
  if (context.textAlign === "center") start = x - width / 2;
  if (context.textAlign === "right") start = x - width;
  const strikeY = y - Math.max(3, Math.min(lineHeight * 0.34, lineHeight - 2));
  context.save();
  context.strokeStyle = context.fillStyle;
  context.lineWidth = 1;
  context.beginPath();
  context.moveTo(start, strikeY);
  context.lineTo(start + width, strikeY);
  context.stroke();
  context.restore();
}



function canvasLayerBounds(sheet, visibleRange) {
  const left = coordinateToPixels(sheet, "col", visibleRange.left - 1);
  const top = coordinateToPixels(sheet, "row", visibleRange.top - 1);
  const right = coordinateToPixels(sheet, "col", visibleRange.right);
  const bottom = coordinateToPixels(sheet, "row", visibleRange.bottom);
  return {
    left,
    top,
    width: Math.max(1, right - left),
    height: Math.max(1, bottom - top),
  };
}



function canvasDevicePixelRatio() {
  return Math.max(1, Math.min(3, Number(window.devicePixelRatio || 1)));
}



function renderCellHtml(sheet, mergeMaps, sheetId, row, col, visibleRange, renderBands = null) {
  const key = cellKey(row, col);
  if (mergeMaps.covered.has(key)) return "";

  const sourceCell = sheet.cells[key] || null;
  const cell = sheetDisplayCell(sheet, row, col, sourceCell) || { row, col };
  const merge = mergeMaps.master.get(key);
  if (merge && renderBands && !rangeIntersectsRenderBands(merge, renderBands)) return "";
  if (merge && !renderBands && !rangeIntersects(merge, visibleRange)) return "";
  if (!cellNeedsDomOverlay(sheet, mergeMaps, row, col, sourceCell, merge)) return "";
  const rowSpan = merge ? merge.bottom - merge.top + 1 : 1;
  const colSpan = merge ? merge.right - merge.left + 1 : 1;
  const rowModel = sheet.rows[row - 1] || {};
  const formula = cell.formula || getCellFormula(sheetId, row, col);
  const display = state.showFormulas && formula ? formula : getDisplayForCell(sheet, sheetId, row, col);
  const overflowWidth = textOverflowWidth(sheet, mergeMaps, row, col, colSpan, cell, display);
  const css = overlayCellCss(cellEffectiveCss(sheet, cell), {
    gridColumn: `${visualGridColumnStart(sheet, col, colSpan)} / span ${colSpan}`,
    gridRow: `${row + 1} / span ${rowSpan}`,
  }, canvasFontDisplayScale(sheet));
  Object.assign(css, frozenCellStickyCss(sheet, row, col));
  const frozenClasses = frozenCellClassNames(sheet, row, col);
  const column = sheet.columns[col - 1];
  if (column?.hidden || rowModel.hidden) css.display = "none";
  const attrs = [
    `data-sheet-index="${state.activeSheetIndex}"`,
    `data-row="${row}"`,
    `data-col="${col}"`,
    `data-address="${columnName(col)}${row}"`,
    `style="${escapeAttr(styleObjectToString(css))}"`,
  ];

  const annotationTitle = cellAnnotationTitle(cell);
  if (annotationTitle) attrs.push(`title="${escapeAttr(annotationTitle)}"`);

  const classNames = ["sheet-cell"];
  if (overflowWidth) classNames.push("text-overflow");
  if (cellHasFill(cell)) classNames.push("has-fill");
  if (merge) classNames.push("is-merged-cell");
  const hideGridRight = shouldHideGridRight(sheet, mergeMaps, row, col, rowSpan, colSpan, cell);
  const hideGridBottom = shouldHideGridBottom(sheet, mergeMaps, row, col, rowSpan, colSpan, cell);
  if (hideGridRight) classNames.push("hide-grid-right");
  if (hideGridBottom) classNames.push("hide-grid-bottom");
  if (cellHasComment(cell)) classNames.push("has-comment");
  if (cellHasNote(cell)) classNames.push("has-note");
  if (cellHasPhonetics(cell)) classNames.push("has-phonetic");
  if (isAutoFilterHeaderCell(sheet, row, col)) classNames.push("has-filter-button");
  if (cell?.hyperlink) classNames.push("has-hyperlink");
  if (cell?.sparkline) classNames.push("has-sparkline");
  if (cellShowsValidationDropdown(cell, sheet)) classNames.push("has-validation-dropdown");
  if (cell?.showPhonetic) classNames.push("show-phonetic");
  if (hasValidationInvalidMark(state.activeSheetIndex, row, col)) classNames.push("has-validation-invalid");
  if (formula) classNames.push("formula-cell");
  classNames.push(...frozenClasses);
  classNames.push(...formulaAuditClassForCell(sheet, row, col));
  if (isCellInActiveSelection(row, col)) {
    classNames.push("is-in-range");
  }
  const formulaReferenceEntry = activeFormulaReferenceEntryForCell(row, col);
  if (formulaReferenceEntry) {
    classNames.push("is-formula-reference");
    classNames.push(formulaReferenceColorClass(formulaReferenceEntry.colorIndex));
  }
  if (isActiveCell(row, col)) {
    classNames.push("is-selected");
  }
  if (classNames.length) attrs.push(`class="${classNames.join(" ")}"`);

  return `<div ${attrs.join(" ")}>${cellInnerHtml(cell, display, overflowWidth, sheet)}</div>`;
}



function cellNeedsDomOverlay(sheet, mergeMaps, row, col, cell, merge = null) {
  if (merge) return true;
  if (isFrozenPaneCell(sheet, row, col)) return true;
  if (isActiveCell(row, col)) return true;
  if (
    state.formulaEdit?.sheetIndex === state.activeSheetIndex &&
    state.formulaEdit.row === row &&
    state.formulaEdit.col === col
  ) {
    return true;
  }
  if (isAutoFilterHeaderCell(sheet, row, col)) return true;
  if (hasValidationInvalidMark(state.activeSheetIndex, row, col)) return true;
  if (activeFormulaReferenceEntryForCell(row, col)) return true;
  if (formulaAuditClassForCell(sheet, row, col).length) return true;
  if (!cell) return false;
  if (cellHasDomContent(cell)) return true;
  if (cell.hyperlink) return true;
  if (cellHasComment(cell) || cellHasNote(cell)) return true;
  if (cellShowsValidationDropdown(cell, sheet)) return true;
  if (cell.sparkline) return true;
  if (cellHasPhonetics(cell) && cell.showPhonetic) return true;
  return false;
}



function cellHasDomContent(cell) {
  if (!cell) return false;
  if (cell.formula || cell.dataTableFormula) return true;
  if (cell.richText?.length) return true;
  if (cell.raw != null && String(cell.raw) !== "") return true;
  if (cell.display != null && String(cell.display) !== "") return true;
  if (cell.cached != null && String(cell.cached) !== "") return true;
  return false;
}



function overlayCellCss(effectiveCss, baseCss, fontScale = 1) {
  const css = { ...baseCss, ...(effectiveCss || {}) };
  css.fontSize = `${canvasFontSize(css.fontSize || "13px", fontScale)}px`;
  if (css.backgroundColor) css["--cell-fill-color"] = css.backgroundColor;
  delete css.backgroundColor;
  delete css.borderColor;
  delete css.borderTopColor;
  delete css.borderRightColor;
  delete css.borderBottomColor;
  delete css.borderLeftColor;
  return css;
}



function cellHasFill(cell) {
  return Boolean(cell?.hasFill || cell?.css?.backgroundColor);
}



function cellHasFontColor(cell) {
  if (cell?.hasFontColor) return true;
  const color = normalizeColorValue(cell?.css?.color);
  return Boolean(color && color !== "#000000");
}



function shouldHideGridRight(sheet, mergeMaps, row, col, rowSpan, colSpan, cell) {
  if (cellHasFill(cell)) return true;
  const rightCol = col + colSpan;
  for (let edgeRow = row; edgeRow < row + rowSpan; edgeRow += 1) {
    if (gridPositionHasFill(sheet, mergeMaps, edgeRow, rightCol)) return true;
  }
  return false;
}



function shouldHideGridBottom(sheet, mergeMaps, row, col, rowSpan, colSpan, cell) {
  if (cellHasFill(cell)) return true;
  const bottomRow = row + rowSpan;
  for (let edgeCol = col; edgeCol < col + colSpan; edgeCol += 1) {
    if (gridPositionHasFill(sheet, mergeMaps, bottomRow, edgeCol)) return true;
  }
  return false;
}



function gridPositionHasFill(sheet, mergeMaps, row, col) {
  if (row < 1 || col < 1 || row > sheet.rowCount || col > sheet.colCount) return false;
  const key = cellKey(row, col);
  const coveredMasterKey = mergeMaps.coveredToMaster.get(key);
  const masterKey = coveredMasterKey || key;
  const [masterRow, masterCol] = coveredMasterKey ? parseCellKey(masterKey) : [row, col];
  return cellHasFill(sheetDisplayCell(sheet, masterRow, masterCol, sheet.cells[masterKey] || null));
}



function worksheetObjectLayerIndex(sheet, image) {
  const images = sheet?.images || [];
  const index = images.findIndex((item) => item?.id === image?.id);
  return index >= 0 ? index : images.length;
}



function worksheetObjectLayerZIndex(sheet, image) {
  return WORKSHEET_OBJECT_LAYER_BASE_Z + worksheetObjectLayerIndex(sheet, image) * WORKSHEET_OBJECT_LAYER_STEP_Z;
}



function worksheetObjectLayerStyle(sheet, image, offset = 0) {
  return `;z-index:${worksheetObjectLayerZIndex(sheet, image) + offset}`;
}



function worksheetGroupSelectionLayerStyle(sheet, images = []) {
  const maxZ = Math.max(...images.map((image) => worksheetObjectLayerZIndex(sheet, image)));
  return `;z-index:${Number.isFinite(maxZ) ? maxZ + WORKSHEET_OBJECT_SELECTION_Z_OFFSET : WORKSHEET_OBJECT_LAYER_BASE_Z}`;
}



function renderImageLayer(sheet, visibleRange) {
  if (sheet?.showObjects === false) return "";
  if (!sheet.images?.length) return "";
  const visibleBounds = imageViewportBounds(sheet, visibleRange);
  return sheet.images
    .map((image) => {
      if (image.hidden) return "";
      const box = imageBoxToPixels(sheet, image);
      if (!box) return "";
      const visualBounds = imageVisualBoundsToPixels(image, box);
      if (!isActiveImage(image.id) && !pixelBoxesIntersect(visualBounds, visibleBounds)) return "";
      const style = imageBoxStyle(image, box, sheet);
      const selected = isActiveImage(image.id) ? " is-selected" : "";
      const imageHtml = imageObjectHtml(image, style, selected, "", sheet.displayScale);
      if (!image.hyperlink) return imageHtml;
      const linkClass = `sheet-image-link${isLineShapeObject(image) ? " is-line-shape" : ""}${selected}`;
      return `<a class="${linkClass}" data-image-id="${escapeAttr(image.id)}" href="${escapeAttr(image.hyperlink)}" target="_blank" rel="noopener noreferrer" style="${escapeAttr(style)}">${imageHtml}</a>`;
    })
    .join("");
}



function renderInkLayer(sheet, visibleRange) {
  if (sheet?.showObjects === false) return "";
  const strokes = normalizeInkStrokes(sheet?.inkStrokes);
  if (!strokes.length) return "";
  const visibleBounds = imageViewportBounds(sheet, visibleRange);
  const selectedIds = selectedInkStrokeIdSet();
  const paths = strokes
    .filter((stroke) => pixelBoxesIntersect(inkStrokeBoundingBox(stroke), visibleBounds))
    .map((stroke) => inkStrokeSvg(stroke, selectedIds.has(stroke.id)))
    .filter(Boolean)
    .join("");
  if (!paths) return "";
  const style = [
    `left:${roundCssNumber(visibleBounds.left)}px`,
    `top:${roundCssNumber(visibleBounds.top)}px`,
    `width:${roundCssNumber(visibleBounds.width)}px`,
    `height:${roundCssNumber(visibleBounds.height)}px`,
  ].join(";");
  return `<svg class="sheet-ink-layer" aria-hidden="true" style="${escapeAttr(style)}" viewBox="${roundCssNumber(visibleBounds.left)} ${roundCssNumber(visibleBounds.top)} ${roundCssNumber(visibleBounds.width)} ${roundCssNumber(visibleBounds.height)}">${paths}</svg>`;
}



function inkStrokeSvg(stroke, selected = false) {
  const d = svgPathDFromSheetPoints(stroke.points || []);
  if (!d) return "";
  const color = stroke.color || "#111827";
  const width = Math.max(0.5, Number(stroke.width) || 2);
  const opacity = Math.max(0.05, Math.min(1, Number(stroke.opacity ?? 1)));
  const selectedPath = selected
    ? `<path class="sheet-ink-stroke-selection" d="${escapeAttr(d)}" fill="none" stroke="#22c55e" stroke-width="${roundCssNumber(width + 7)}" stroke-linecap="round" stroke-linejoin="round" opacity="0.22"></path>`
    : "";
  return `${selectedPath}<path class="sheet-ink-stroke${selected ? " is-selected" : ""}" data-ink-stroke-id="${escapeAttr(stroke.id)}" d="${escapeAttr(d)}" fill="none" stroke="${escapeAttr(color)}" stroke-width="${roundCssNumber(width)}" stroke-linecap="round" stroke-linejoin="round" opacity="${roundCssNumber(opacity)}"></path>`;
}



function imageObjectHtml(image, style, selected = "", extraAttrs = "", textDisplayScale = 1) {
  const crop = normalizePictureCrop(image?.crop);
  const isLineShape = isLineShapeObject(image);
  const imageClass = `sheet-image${image.shape ? " sheet-shape" : ""}${isLineShape ? " is-line-shape" : ""}${crop ? " sheet-picture-crop" : ""}${selected}`;
  const alt = escapeAttr(image.name || "worksheet image");
  const attrs = extraAttrs ? ` ${extraAttrs}` : "";
  const src = worksheetImageRenderSource(image);
  if (image?.shape && !isLineShape && !crop) {
    return `<div class="${imageClass}" data-image-id="${escapeAttr(image.id)}"${attrs} style="${escapeAttr(style)}" role="img" aria-label="${alt}"><img class="sheet-shape-render" src="${escapeAttr(src)}" alt="" draggable="false" loading="lazy" decoding="async" />${shapeTextLayerHtml(image, textDisplayScale)}</div>`;
  }
  if (!crop) {
    return `<img class="${imageClass}" data-image-id="${escapeAttr(image.id)}"${attrs} src="${escapeAttr(src)}" alt="${alt}" style="${escapeAttr(style)}" draggable="false" loading="lazy" decoding="async" />`;
  }
  return `<div class="${imageClass}" data-image-id="${escapeAttr(image.id)}"${attrs} style="${escapeAttr(style)}" role="img" aria-label="${alt}"><img class="sheet-picture-crop-content" src="${escapeAttr(src)}" alt="" style="${escapeAttr(pictureCropContentStyle(crop))}" draggable="false" loading="lazy" decoding="async" /></div>`;
}



function shapeTextLayerHtml(image, textDisplayScale = 1) {
  const shape = image?.shape;
  if (!shape || isLineShapeObject(image)) return "";
  const text = shapeTextPlainValue(shape);
  const runs = normalizeShapeTextRuns(shape.textRuns);
  if (!text && !runs.length) return "";
  const content = runs.length ? shapeTextRunsLayerHtml(runs) : shapeTextPlainLayerHtml(text);
  return `<div class="sheet-shape-text-layer" style="${escapeAttr(shapeTextLayerStyle(shape) + ";" + shapeTextLayerScaleStyle(textDisplayScale))}"><div class="sheet-shape-text-block" style="width:100%;min-width:0;flex:none">${content}</div></div>`;
}



function shapeTextLayerScaleStyle(value) {
  const scale = Number(value);
  if (!Number.isFinite(scale) || scale <= 0 || scale === 1) return "";
  // Lay out plain/rich text at its original size, then scale the entire layer
  // with the shape. This keeps wrapping, padding and text outlines consistent.
  return `width:${100 / scale}%;height:${100 / scale}%;transform-origin:0 0;transform:scale(${scale})`;
}



function shapeTextPlainValue(shape) {
  if (!shape) return "";
  if (Object.prototype.hasOwnProperty.call(shape, "text")) return String(shape.text ?? "");
  return shapePlainTextFromRuns(shape.textRuns) || "";
}



function shapeTextLayerStyle(shape = {}) {
  const margin = normalizeShapeTextMargin(shape.textMargin, shape.type === "textBox");
  const vertical = ["top", "bottom", "middle"].includes(shape.textVAlign) ? shape.textVAlign : shape.type === "textBox" ? "top" : "middle";
  const justifyContent = vertical === "top" ? "flex-start" : vertical === "bottom" ? "flex-end" : "center";
  const underline = underlinePresetForValue(shape.textUnderline)?.value || "";
  const fontFamily = shape.textFontFamily || fontFamilyForExcel("Calibri") || '"Yu Gothic UI", Meiryo, sans-serif';
  const fontSize = Math.max(8, Math.round((Number(shape.textFontSize) || 11) * 96 / 72));
  const styles = [
    "position:absolute;inset:0;box-sizing:border-box;display:flex;flex-direction:column;min-width:0;min-height:0;overflow:hidden;line-height:1.22;white-space:pre-wrap;overflow-wrap:break-word;word-break:normal",
    `padding:${roundCssNumber(margin.top)}px ${roundCssNumber(margin.right)}px ${roundCssNumber(margin.bottom)}px ${roundCssNumber(margin.left)}px`,
    `color:${shape.textColor || "#1f2937"}`,
    `font-family:${fontFamily}`,
    `font-size:${fontSize}px`,
    `font-weight:${shape.textBold ? "700" : "400"}`,
    `font-style:${shape.textItalic ? "italic" : "normal"}`,
    `text-align:${shape.textAlign || "center"}`,
    `justify-content:${justifyContent}`,
    underline ? "text-decoration-line:underline" : "text-decoration-line:none",
    underline === "underline double" ? "text-decoration-style:double" : "",
    shape.textOutlineColor && shape.textOutlineColor !== "none" ? `-webkit-text-stroke:${normalizeShapeTextOutlineWidth(shape.textOutlineWidth)}px ${shape.textOutlineColor}` : "",
  ].filter(Boolean);
  return styles.join(";");
}



function shapeTextPlainLayerHtml(text) {
  return splitShapeTextLines(text).map((line) => (line ? escapeHtml(line) : "&#8203;")).join("<br>");
}



function shapeTextRunsLayerHtml(runs) {
  const html = normalizeShapeTextRuns(runs)
    .map((run) => {
      const style = shapeTextRunLayerStyle(run);
      const attrs = style ? ` style="${escapeAttr(style)}"` : "";
      return splitShapeTextLines(run.text)
        .map((line) => `<span${attrs}>${line ? escapeHtml(line) : "&#8203;"}</span>`)
        .join("<br>");
    })
    .join("");
  return html || "&#8203;";
}



function splitShapeTextLines(text) {
  return String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
}



function shapeTextRunLayerStyle(run = {}) {
  const underline = underlinePresetForValue(run.underline)?.value || "";
  return [
    run.color ? `color:${run.color}` : "",
    Number.isFinite(Number(run.fontSize)) ? `font-size:${Math.max(8, Math.round(Number(run.fontSize) * 96 / 72))}px` : "",
    run.fontFamily ? `font-family:${run.fontFamily}` : "",
    typeof run.bold === "boolean" ? `font-weight:${run.bold ? "700" : "400"}` : "",
    typeof run.italic === "boolean" ? `font-style:${run.italic ? "italic" : "normal"}` : "",
    underline ? "text-decoration-line:underline" : "",
    underline === "underline double" ? "text-decoration-style:double" : "",
  ].filter(Boolean).join(";");
}



function setImageElementRenderSource(element, src) {
  if (!element || !src) return;
  if (element.matches?.("img.sheet-image")) {
    element.src = src;
    return;
  }
  const target = element.querySelector?.(".sheet-shape-render, .sheet-picture-crop-content, img.sheet-image");
  if (target) target.setAttribute("src", src);
}



function worksheetImageRenderSource(image) {
  if (image?.chart?.series?.length && (!image.shape || image.shape.sourcePreset === "graphicFrame")) {
    const width = Math.max(1, Math.round(Number(image.ext?.width) || 520));
    const height = Math.max(1, Math.round(Number(image.ext?.height) || 320));
    return svgDataUrl(chartSvg(image.chart, image.chart.type || "column", { width, height, title: image.chart.title || image.name || "グラフ" }));
  }
  if (image?.shape) {
    const descriptor = shapeDescriptorForType(image.shape.type || DEFAULT_SHAPE_TYPE, {
      ...image.shape,
      width: image.ext?.width,
      height: image.ext?.height,
      includeTextInSvg: false,
    });
    return cachedWorksheetImageRenderSource(descriptor.src || image.src || "");
  }
  return cachedWorksheetImageRenderSource(image?.src || "");
}



function renderImageSelectionFrames(sheet, visibleRange) {
  if (sheet?.showObjects === false) return "";
  const images = selectedImageObjects().filter((image) => !image.hidden);
  const renderedIds = new Set();
  const frames = [];
  selectedImageGroups(images).forEach((group) => {
    const frame = renderImageGroupSelectionFrame(sheet, visibleRange, group);
    if (!frame) return;
    frames.push(frame);
    group.images.forEach((image) => renderedIds.add(image.id));
  });
  images.forEach((image) => {
    if (renderedIds.has(image.id)) return;
    frames.push(renderImageSelectionFrame(sheet, visibleRange, image));
  });
  return frames.join("");
}



function selectedCroppablePicture() {
  const objects = selectedImageObjects();
  if (objects.length !== 1) return null;
  const image = objects[0];
  return image && !image.shape && !image.chart ? image : null;
}



function isPictureCropModeActive(image = selectedCroppablePicture()) {
  return Boolean(image && state.pictureCrop?.sheetIndex === state.activeSheetIndex && state.pictureCrop?.imageId === image.id);
}



function clearPictureCropState() {
  const hadCrop = Boolean(state.pictureCrop || state.pictureCropDrag);
  state.pictureCrop = null;
  state.pictureCropDrag = null;
  return hadCrop;
}



function syncPictureCropSelectionState() {
  const image = selectedCroppablePicture();
  if (image && isPictureCropModeActive(image)) return false;
  return clearPictureCropState();
}



function renderPictureCropHandles() {
  return ["nw", "n", "ne", "e", "se", "s", "sw", "w"]
    .map((handle) => `<span class="sheet-picture-crop-handle is-${handle}" data-picture-crop-handle="${handle}" aria-hidden="true"></span>`)
    .join("");
}



function renderImageSelectionFrame(sheet, visibleRange, image) {
  const box = imageBoxToPixels(sheet, image);
  if (!box || !pixelBoxesIntersect(box, imageViewportBounds(sheet, visibleRange))) return "";
  const style = imageSelectionFrameStyle(image, box, sheet);
  const isLine = isLineShapeObject(image);
  const isCopySource = !isLine && isCopiedSourceImage(image.id);
  const isPointEdit = (state.shapePointEditIds || []).includes(image.id);
  const isCropMode = isPictureCropModeActive(image);
  const frameClass = `sheet-image-selection-frame${image.shape ? " is-shape" : ""}${isLine ? " is-line-shape" : ""}${isCopySource ? " is-copy-source" : ""}${isPointEdit ? " is-editing-points" : ""}${isCropMode ? " is-picture-crop-mode" : ""}`;
  const handles = isLine ? renderLineShapeHandles(image) : isCropMode ? renderPictureCropHandles() : renderShapeResizeHandles(image);
  return `<div class="${frameClass}" data-image-id="${escapeAttr(image.id)}" style="${escapeAttr(style)}" aria-hidden="true">${handles}</div>`;
}



function renderImageGroupSelectionFrame(sheet, visibleRange, group) {
  const box = unionImageSelectionBounds(sheet, group.images);
  if (!box || !pixelBoxesIntersect(box, imageViewportBounds(sheet, visibleRange))) return "";
  const style = `left:${roundCssNumber(box.left)}px;top:${roundCssNumber(box.top)}px;width:${roundCssNumber(box.width)}px;height:${roundCssNumber(box.height)}px${worksheetGroupSelectionLayerStyle(sheet, group.images)}`;
  const isCopySource = group.images.some((image) => isCopiedSourceImage(image.id));
  const frameClass = `sheet-image-selection-frame is-shape is-group-shape${isCopySource ? " is-copy-source" : ""}`;
  return `<div class="${frameClass}" data-image-id="${escapeAttr(group.primaryId)}" data-image-group-id="${escapeAttr(group.groupId)}" style="${escapeAttr(style)}" aria-hidden="true">${renderGroupShapeHandles()}</div>`;
}



function selectedImageGroups(images = selectedImageObjects()) {
  const byGroup = new Map();
  images.forEach((image) => {
    if (!image?.groupId) return;
    const group = byGroup.get(image.groupId) || { groupId: image.groupId, primaryId: image.id, images: [] };
    group.images.push(image);
    byGroup.set(image.groupId, group);
  });
  return [...byGroup.values()].filter((group) => group.images.length > 1);
}



function isCopiedSourceImage(imageId) {
  const clipboard = state.objectClipboard;
  return Boolean(
    imageId &&
      clipboard?.type === "objects" &&
      clipboard.sourceSheetIndex === state.activeSheetIndex &&
      (clipboard.sourceIds || []).includes(imageId),
  );
}



function imageBoxStyle(image, box, sheet = activeSheet()) {
  return imageBoxBaseStyle(image, imageVisualBoxToPixels(image, box)) + objectVisualStyle(image) + worksheetObjectLayerStyle(sheet, image);
}



function imageVisualBoxToPixels(image, box) {
  if (!box || !isCalloutShapeType(image?.shape?.type) || imageRotationDegrees(image)) return box;
  const { viewBox } = calloutGeometry(normalizeShapeAdjustments(image.shape.type, image.shape.adjust));
  const scaleX = box.width / 100;
  const scaleY = box.height / 60;
  return {
    left: box.left + viewBox.left * scaleX,
    top: box.top + viewBox.top * scaleY,
    width: Math.max(1, viewBox.width * scaleX),
    height: Math.max(1, viewBox.height * scaleY),
  };
}



function imageVisualBoundsToPixels(image, box, rotationOverride = undefined) {
  const visualBox = imageVisualBoxToPixels(image, box);
  return rotatedPixelBox(visualBox, rotationOverride === undefined ? imageRotationDegrees(image) : rotationOverride);
}



function imageSelectionBoundsToPixels(image, box, rotationOverride = undefined) {
  const selectionBox = imageSelectionBoxToPixels(image, box);
  return rotatedPixelBox(selectionBox, rotationOverride === undefined ? imageRotationDegrees(image) : rotationOverride);
}



function imageSelectionBoxToPixels(image, box) {
  if (!box || !isCalloutShapeType(image?.shape?.type) || imageRotationDegrees(image)) return box;
  const bounds = calloutBodySelectionBounds(image.shape.type);
  const scaleX = box.width / 100;
  const scaleY = box.height / 60;
  return {
    left: box.left + bounds.left * scaleX,
    top: box.top + bounds.top * scaleY,
    width: Math.max(1, (bounds.right - bounds.left) * scaleX),
    height: Math.max(1, (bounds.bottom - bounds.top) * scaleY),
  };
}



function rotatedPixelBox(box, degrees = 0) {
  if (!box) return null;
  const rotation = normalizeRotationDegrees(degrees);
  if (!rotation) return box;
  const center = {
    x: box.left + box.width / 2,
    y: box.top + box.height / 2,
  };
  const points = [
    { x: box.left, y: box.top },
    { x: box.left + box.width, y: box.top },
    { x: box.left + box.width, y: box.top + box.height },
    { x: box.left, y: box.top + box.height },
  ].map((point) => rotatePointAround(point, center, rotation));
  const left = Math.min(...points.map((point) => point.x));
  const top = Math.min(...points.map((point) => point.y));
  const right = Math.max(...points.map((point) => point.x));
  const bottom = Math.max(...points.map((point) => point.y));
  return {
    left,
    top,
    width: Math.max(1, right - left),
    height: Math.max(1, bottom - top),
  };
}



function imageBoxBaseStyle(image, box) {
  const rotation = imageRotationDegrees(image);
  const transforms = [];
  if (rotation) transforms.push(`rotate(${roundCssNumber(rotation)}deg)`);
  if (imageFlipH(image)) transforms.push("scaleX(-1)");
  if (imageFlipV(image)) transforms.push("scaleY(-1)");
  const transform = transforms.length ? `;transform:${transforms.join(" ")};transform-origin:center center` : "";
  return `left:${roundCssNumber(box.left)}px;top:${roundCssNumber(box.top)}px;width:${roundCssNumber(box.width)}px;height:${roundCssNumber(box.height)}px${transform}`;
}



function imageSelectionFrameStyle(image, box, sheet = activeSheet()) {
  return imageBoxBaseStyle(image, imageSelectionBoxToPixels(image, box)) + worksheetObjectLayerStyle(sheet, image, WORKSHEET_OBJECT_SELECTION_Z_OFFSET);
}



function objectVisualStyle(image) {
  return `${chartObjectOutlineStyle(image)}${pictureOutlineStyle(image)}${shapeVisualEffectStyle(image?.shape || image, image?.src || "", image)}`;
}



function chartObjectOutlineStyle(image) {
  if (!image?.shape || !image?.chart?.series?.length) return "";
  const color = normalizeColorValue(image.shape.outlineColor);
  if (!color || color === "none") return "";
  const width = normalizeShapeOutlineWidth(image.shape.outlineWidth);
  return `;border:${roundCssNumber(width)}px ${pictureOutlineCssStyle(image.shape)} ${color}`;
}



function pictureOutlineStyle(image) {
  if (!image || image.shape) return "";
  const color = normalizeColorValue(image.outlineColor);
  if (!color || color === "none") return "";
  const width = normalizeShapeOutlineWidth(image.outlineWidth);
  return `;border:${roundCssNumber(width)}px ${pictureOutlineCssStyle(image)} ${color}`;
}



function pictureOutlineCssStyle(image) {
  const compound = normalizeShapeLineCompound(image?.lineCompound);
  if (compound && compound !== "single") return "double";
  const dash = normalizeShapeLineDash(image?.lineDash);
  if (dash === "dot" || dash === "squareDot") return "dotted";
  if (dash && dash !== "solid") return "dashed";
  return "solid";
}



function pictureCropContentStyle(crop) {
  const normalized = normalizePictureCrop(crop) || { left: 0, top: 0, right: 0, bottom: 0 };
  const visibleWidth = Math.max(0.01, 1 - normalized.left - normalized.right);
  const visibleHeight = Math.max(0.01, 1 - normalized.top - normalized.bottom);
  return [
    "position:absolute",
    "display:block",
    `left:${roundCssNumber((-normalized.left / visibleWidth) * 100)}%`,
    `top:${roundCssNumber((-normalized.top / visibleHeight) * 100)}%`,
    `width:${roundCssNumber((1 / visibleWidth) * 100)}%`,
    `height:${roundCssNumber((1 / visibleHeight) * 100)}%`,
    "max-width:none",
    "max-height:none",
    "pointer-events:none",
    "user-select:none",
  ].join(";");
}



function shapeVisualEffectStyle(shape, source = "", image = null) {
  const isChartGraphicFrame = Boolean(image?.chart?.series?.length && shape?.sourcePreset === "graphicFrame");
  const hasEmbeddedSvgEffects = Boolean(shape) && String(source || "").startsWith("data:image/svg+xml") && !isChartGraphicFrame;
  const filters = hasEmbeddedSvgEffects ? [] : shapeFilterValues(shape);
  const filter = filters.length ? `;filter:${filters.join(" ")}` : "";
  const reflection = normalizeShapeReflection(shape?.reflection);
  const reflectionStyle = reflection ? `;-webkit-box-reflect:below ${roundCssNumber(reflection.distance || 0)}px linear-gradient(transparent, rgba(255,255,255,${roundCssNumber(reflection.alpha || 0.45)}))` : "";
  return `${filter}${reflectionStyle}`;
}



function shapeFilterValues(shape) {
  const filters = [];
  if (shape?.shadow && typeof shape.shadow === "object") {
    const dx = roundCssNumber(Number(shape.shadow.dx) || 0);
    const dy = roundCssNumber(Number(shape.shadow.dy) || 0);
    const blur = roundCssNumber(Math.max(0, Number(shape.shadow.blur) || 0));
    const color = cssColorForSvg(shape.shadow.color) || "rgba(0,0,0,0.35)";
    filters.push(`drop-shadow(${dx}px ${dy}px ${blur}px ${color})`);
  }
  if (shape?.shadow === "outer1") {
    filters.push("drop-shadow(2px 2px 2px rgba(0,0,0,0.28))");
  } else if (shape?.shadow === "outer2") {
    filters.push("drop-shadow(3px 4px 5px rgba(0,0,0,0.24))");
  }
  const glow = normalizeShapeGlow(shape?.glow);
  if (glow) filters.push(`drop-shadow(0 0 ${roundCssNumber(glow.radius)}px ${cssColorForSvg(glow.color) || "#5b9bd5"})`);
  const softEdges = normalizeShapeSoftEdges(shape?.softEdges);
  if (softEdges) filters.push(`blur(${roundCssNumber(softEdges * 0.55)}px)`);
  return filters;
}



function renderShapeResizeHandles(image) {
  const adjustHandles = renderShapeAdjustmentHandles(image);
  const rotateHandle = image?.shape
    ? '<span class="sheet-image-rotate-stem" aria-hidden="true"></span><span class="sheet-image-rotate-handle" data-image-rotate="true" aria-hidden="true"></span>'
    : image
      ? '<span class="sheet-image-rotate-stem" aria-hidden="true"></span><span class="sheet-image-rotate-handle" data-image-rotate="true" aria-hidden="true"></span>'
    : "";
  return `${renderResizeHandles()}${adjustHandles}${rotateHandle}`;
}



function renderGroupShapeHandles() {
  return `${renderResizeHandles()}<span class="sheet-image-rotate-stem" aria-hidden="true"></span><span class="sheet-image-rotate-handle" data-image-rotate="true" aria-hidden="true"></span>`;
}



function renderResizeHandles() {
  return ["nw", "n", "ne", "e", "se", "s", "sw", "w"]
    .map((handle) => `<span class="sheet-image-resize-handle is-${handle}" data-image-resize="${handle}" aria-hidden="true"></span>`)
    .join("");
}



function renderShapeAdjustmentHandles(image) {
  if (!image?.shape || activeSelectedImageIds().length !== 1) return "";
  if (image.shape.type === CUSTOM_SHAPE_TYPE) {
    if (!(state.shapePointEditIds || []).includes(image.id)) return "";
    return shapeCustomPathHandles(image.shape.customPath)
      .map((handle) => `<span class="sheet-shape-adjust-handle is-vertex" data-shape-adjust="${escapeAttr(handle.key)}" style="left:${roundCssNumber(handle.x * 100)}%;top:${roundCssNumber(handle.y * 100)}%" aria-hidden="true"></span>`)
      .join("");
  }
  return shapeAdjustmentHandlesForFrame(image.shape)
    .map((handle) => `<span class="sheet-shape-adjust-handle" data-shape-adjust="${escapeAttr(handle.key)}" style="left:${roundCssNumber(handle.x * 100)}%;top:${roundCssNumber(handle.y * 100)}%" aria-hidden="true"></span>`)
    .join("");
}



function shapeAdjustmentHandlesForFrame(shape, box = null) {
  if (shape?.type === "elbowConnector") {
    return elbowConnectorAdjustmentHandlesForFrame(shape, box || shapeAdjustmentBoxFromShape(shape));
  }
  const handles = shapeAdjustmentHandles(shape?.type, shape?.adjust);
  if (shape?.type === "roundRect") {
    const size = box || shapeAdjustmentBoxFromShape(shape);
    const scale = roundRectCornerScale(size.width, size.height);
    return handles.map(handle => ({ ...handle, x: handle.x * scale.x }));
  }
  if (!isCalloutShapeType(shape?.type)) return handles;
  const body = calloutBodySelectionBounds(shape.type);
  const bodyWidth = Math.max(0.01, body.right - body.left);
  const bodyHeight = Math.max(0.01, body.bottom - body.top);
  return handles.map((handle) => ({
    ...handle,
    x: ((handle.x * 100) - body.left) / bodyWidth,
    y: ((handle.y * 60) - body.top) / bodyHeight,
  }));
}



function elbowConnectorAdjustmentHandlesForFrame(shape, box = {}) {
  const normalized = normalizeShapeAdjustments("elbowConnector", shape?.adjust) || { bend: 0.5 };
  const points = normalizeShapeLinePoints(shape?.points);
  const horizontalMiddle = elbowConnectorUsesHorizontalMiddle(shape, box);
  const bend = clampElbowConnectorBendForPoints(points, normalized.bend, horizontalMiddle);
  if (horizontalMiddle) {
    return [{ key: "bend", x: (points.x1 + points.x2) / 2, y: bend }];
  }
  return [{ key: "bend", x: bend, y: (points.y1 + points.y2) / 2 }];
}



function elbowConnectorUsesHorizontalMiddle(shape, box = {}) {
  const points = normalizeShapeLinePoints(shape?.points);
  const width = Math.max(1, Number(box?.width ?? shape?.width) || 1);
  const height = Math.max(1, Number(box?.height ?? shape?.height) || 1);
  const dx = Math.abs(points.x2 - points.x1) * width;
  const dy = Math.abs(points.y2 - points.y1) * height;
  return shape?.adjust?.axis === "y" || (shape?.adjust?.axis !== "x" && dy > dx * 1.2);
}



function clampElbowConnectorBendForPoints(points, bend, horizontalMiddle) {
  const normalized = normalizeShapeLinePoints(points);
  const a = horizontalMiddle ? normalized.y1 : normalized.x1;
  const b = horizontalMiddle ? normalized.y2 : normalized.x2;
  const min = Math.min(a, b);
  const max = Math.max(a, b);
  return clampNumber(bend, min, max, (a + b) / 2);
}



function shapeAdjustmentBoxFromShape(shape) {
  return {
    width: Math.max(1, Number(shape?.width) || 1),
    height: Math.max(1, Number(shape?.height) || 1),
  };
}



function shapeAdjustmentBoxFromFrame(frame) {
  const rect = frame?.getBoundingClientRect?.();
  return {
    width: Math.max(1, Number(rect?.width) || 1),
    height: Math.max(1, Number(rect?.height) || 1),
  };
}



function shapeCustomPathHandles(customPath) {
  const path = normalizeShapeCustomPath(customPath);
  if (!path) return [];
  return path.commands.flatMap((command, commandIndex) =>
    command.points.map((point, pointIndex) => ({
      key: `custom-point:${commandIndex}:${pointIndex}`,
      x: point.x,
      y: point.y,
    })),
  );
}



function parseCustomPathVertexKey(vertexKey) {
  const match = String(vertexKey || "").match(/^custom-point:(\d+):(\d+)$/);
  if (!match) return null;
  return { commandIndex: Number(match[1]), pointIndex: Number(match[2]) };
}



function customPathVertexForKey(customPath, vertexKey) {
  const path = normalizeShapeCustomPath(customPath);
  const parsed = parseCustomPathVertexKey(vertexKey);
  if (!path || !parsed) return null;
  const command = path.commands[parsed.commandIndex];
  const point = command?.points?.[parsed.pointIndex];
  return command && point ? { path, command, point, ...parsed } : null;
}



function customPathEditablePointCount(customPath) {
  return normalizeShapeCustomPath(customPath)?.commands.reduce((count, command) => count + command.points.length, 0) || 0;
}



function shapeCustomPathIsClosed(customPath) {
  const commands = normalizeShapeCustomPath(customPath)?.commands || [];
  return commands.some((command) => command.cmd === "Z");
}



function shapeAdjustmentHandles(type, adjust = {}) {
  const normalized = normalizeShapeAdjustments(type, adjust);
  if (!normalized) return [];
  switch (type) {
    case "elbowConnector":
      return [{ key: "bend", x: normalized.bend, y: 0.5 }];
    case "curvedConnector":
      return [{ key: "curve", x: 0.5, y: normalized.curve }];
    case "roundRect":
      return [{ key: "corner", x: normalized.corner, y: 0 }];
    case "roundSnipRect":
      return [
        { key: "corner", x: normalized.corner, y: 0 },
        { key: "cut", x: 1 - normalized.cut, y: 0 },
      ];
    case "snipRect":
      return [{ key: "cut", x: 1 - normalized.cut, y: 0 }];
    case "parallelogram":
    case "flowData":
      return [{ key: "skew", x: normalized.skew, y: 0 }];
    case "trapezoid":
      return [{ key: "inset", x: normalized.inset, y: 0 }];
    case "cross":
    case "leftRightBlockArrow":
    case "leftRightUpArrow":
    case "quadArrow":
    case "bentUpArrow":
    case "uturnArrow":
    case "curvedRightArrow":
      return [{ key: "thickness", x: 0.5, y: (1 - normalized.thickness) / 2 }];
    case "notchedRightArrow":
      return [
        { key: "thickness", x: 0.5, y: (1 - normalized.thickness) / 2 },
        { key: "head", x: normalized.head, y: 0.5 },
      ];
    case "rightBlockArrow":
    case "leftBlockArrow":
      return [
        { key: "thickness", x: 0.5, y: (1 - normalized.thickness) / 2 },
        { key: "head", x: normalized.head, y: 0.5 },
      ];
    case "upBlockArrow":
    case "downBlockArrow":
      return [
        { key: "thickness", x: (1 - normalized.thickness) / 2, y: 0.5 },
        { key: "head", x: 0.5, y: normalized.head },
      ];
    case "can":
      return [{ key: "depth", x: 0.5, y: normalized.depth }];
    case "flowManualInput":
      return [{ key: "skew", x: 0, y: normalized.skew }];
    case "star4":
    case "star5":
    case "star6":
    case "star7":
    case "star8":
    case "star10":
    case "star12":
    case "star16":
    case "star24":
    case "star32":
      return [{ key: "inner", x: 0.5 + normalized.inner * 0.34, y: 0.5 }];
    case "calloutRect":
    case "calloutRoundRect":
    case "calloutOval":
    case "calloutCloud":
      return [{ key: "tail", x: normalized.tailX, y: normalized.tailY }];
    default:
      return [];
  }
}



function renderLineShapeHandles(image) {
  const points = normalizeShapeLinePoints(image?.shape?.points);
  const endpointHandles = ["start", "end"]
    .map((name) => {
      const x = name === "start" ? points.x1 : points.x2;
      const y = name === "start" ? points.y1 : points.y2;
      return `<span class="sheet-image-resize-handle is-line-${name}" data-image-resize="line-${name}" style="left:${roundCssNumber(x * 100)}%;top:${roundCssNumber(y * 100)}%" aria-hidden="true"></span>`;
    })
    .join("");
  const adjustmentHandles = activeSelectedImageIds().length === 1
    ? shapeAdjustmentHandlesForFrame(
      image?.shape,
      { width: Number(image?.ext?.width) || 1, height: Number(image?.ext?.height) || 1 },
    )
      .map((handle) => `<span class="sheet-shape-adjust-handle" data-shape-adjust="${escapeAttr(handle.key)}" style="left:${roundCssNumber(handle.x * 100)}%;top:${roundCssNumber(handle.y * 100)}%" aria-hidden="true"></span>`)
      .join("")
    : "";
  return `${endpointHandles}${adjustmentHandles}`;
}



function imageRotationDegrees(image) {
  if (!image) return 0;
  return normalizeRotationDegrees(image.shape ? image.shape.rotation : image.rotation);
}



function imageFlipH(image) {
  return Boolean(image?.shape ? image.shape.flipH : image?.flipH);
}



function imageFlipV(image) {
  return Boolean(image?.shape ? image.shape.flipV : image?.flipV);
}



function roundCssNumber(value) {
  return Math.round(Number(value) * 100) / 100;
}



function imageViewportBounds(sheet, visibleRange) {
  const host = $sheetHost[0];
  if (!host) return canvasLayerBounds(sheet, visibleRange);
  return {
    left: Math.max(0, host.scrollLeft - IMAGE_RENDER_BUFFER_PX),
    top: Math.max(0, host.scrollTop - IMAGE_RENDER_BUFFER_PX),
    width: host.clientWidth + IMAGE_RENDER_BUFFER_PX * 2,
    height: host.clientHeight + IMAGE_RENDER_BUFFER_PX * 2,
  };
}



function pixelBoxesIntersect(box, bounds) {
  const boxRight = box.left + box.width;
  const boxBottom = box.top + box.height;
  const boundsRight = bounds.left + bounds.width;
  const boundsBottom = bounds.top + bounds.height;
  return boxRight >= bounds.left && box.left <= boundsRight && boxBottom >= bounds.top && box.top <= boundsBottom;
}



function unionImageBox(sheet, images) {
  const boxes = (images || []).map((image) => imageBoxToPixels(sheet, image)).filter(Boolean);
  if (!boxes.length) return null;
  const left = Math.min(...boxes.map((box) => box.left));
  const top = Math.min(...boxes.map((box) => box.top));
  const right = Math.max(...boxes.map((box) => box.left + box.width));
  const bottom = Math.max(...boxes.map((box) => box.top + box.height));
  return {
    left,
    top,
    width: Math.max(1, right - left),
    height: Math.max(1, bottom - top),
  };
}



function unionImageSelectionBounds(sheet, images, boxById = null, rotationById = null) {
  const boxes = (images || [])
    .map((image) => {
      const box = boxById && Object.prototype.hasOwnProperty.call(boxById, image.id)
        ? boxById[image.id]
        : imageBoxToPixels(sheet, image);
      if (!box) return null;
      const rotation = rotationById && Object.prototype.hasOwnProperty.call(rotationById, image.id)
        ? rotationById[image.id]
        : undefined;
      return imageSelectionBoundsToPixels(image, box, rotation);
    })
    .filter(Boolean);
  return unionPixelBoxes(boxes);
}



function renderSelectionFrame(sheet, visibleRange) {
  const ranges = activeSelectionRanges();
  const isMultiRange = ranges.length > 1;
  return ranges
    .filter((range) => rangeIntersects(range, visibleRange))
    .map((range) => {
      const isCopySource = isCopiedSourceRange(range);
      const className = `sheet-selection-frame${isCopySource ? " is-copy-source" : ""}${isMultiRange && !isCopySource ? " is-multi-range" : ""}`;
      return renderRangeFrameHtml(sheet, range, visibleRange, className);
    })
    .join("");
}



function renderFormulaReferenceFrame(sheet, visibleRange) {
  return activeFormulaReferenceEntriesForActiveSheet()
    .filter((range) => rangeIntersects(range, visibleRange))
    .map((range) => renderRangeFrameHtml(sheet, range, visibleRange, `sheet-formula-reference-frame ${formulaReferenceColorClass(range.colorIndex)}`))
    .join("");
}



function renderCopyFrame(sheet, visibleRange) {
  const range = copiedRangeForActiveSheet();
  if (!range || !rangeIntersects(range, visibleRange) || activeSelectionRanges().some((activeRange) => rangesEqual(range, activeRange))) return "";
  return renderRangeFrameHtml(sheet, range, visibleRange, "sheet-copy-frame");
}



function renderRangeFrameHtml(sheet, range, visibleRange, className) {
  if (!range || !rangeIntersects(range, visibleRange)) return "";
  const frameRange = clipRangeToVisibleRange(range, visibleRange);
  if (!frameRange) return "";
  const left = coordinateToPixels(sheet, "col", frameRange.left - 1);
  const top = coordinateToPixels(sheet, "row", frameRange.top - 1);
  const right = coordinateToPixels(sheet, "col", frameRange.right);
  const bottom = coordinateToPixels(sheet, "row", frameRange.bottom);
  const width = Math.max(1, right - left);
  const height = Math.max(1, bottom - top);
  const style = `left:${left}px;top:${top}px;width:${width}px;height:${height}px;${clippedRangeFrameBorderStyle(range, frameRange)}`;
  const handle = className.includes("sheet-selection-frame") &&
    !className.includes("is-copy-source") &&
    !className.includes("is-multi-range") &&
    !range.fullRow &&
    !range.fullColumn &&
    frameRange.bottom === range.bottom &&
    frameRange.right === range.right
    ? '<span class="sheet-fill-handle" aria-hidden="true"></span>'
    : "";
  return `<div class="${className}" aria-hidden="true" style="${escapeAttr(style)}">${handle}</div>`;
}



function clipRangeToVisibleRange(range, visibleRange) {
  if (!range || !visibleRange || !rangeIntersects(range, visibleRange)) return null;
  return {
    ...range,
    top: Math.max(range.top, visibleRange.top),
    left: Math.max(range.left, visibleRange.left),
    bottom: Math.min(range.bottom, visibleRange.bottom),
    right: Math.min(range.right, visibleRange.right),
  };
}



function clippedRangeFrameBorderStyle(range, frameRange) {
  if (!range || !frameRange) return "";
  return [
    frameRange.top > range.top ? "border-top-color:transparent" : "",
    frameRange.left > range.left ? "border-left-color:transparent" : "",
    frameRange.bottom < range.bottom ? "border-bottom-color:transparent" : "",
    frameRange.right < range.right ? "border-right-color:transparent" : "",
  ].filter(Boolean).join(";");
}



function copiedRangeForActiveSheet() {
  const clipboard = state.internalClipboard;
  if (!clipboard || clipboard.sourceSheetIndex !== state.activeSheetIndex || !clipboard.sourceRange) return null;
  return clipboard.sourceRange;
}



function isCopiedSourceRange(range) {
  return rangesEqual(range, copiedRangeForActiveSheet());
}



function rangesEqual(a, b) {
  if (!a || !b) return false;
  return a.sheetIndex === b.sheetIndex && a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
}



function rawRangesEqual(a, b) {
  return Boolean(
    a &&
      b &&
      rangesEqual(a, b) &&
      Boolean(a.fullRow) === Boolean(b.fullRow) &&
      Boolean(a.fullColumn) === Boolean(b.fullColumn),
  );
}



function isActiveImage(imageId) {
  return activeSelectedImageIds().includes(imageId);
}



function activeSelectedImageIds() {
  if (state.selectedImage?.sheetIndex !== state.activeSheetIndex) return [];
  const ids = (state.selectedImages || [])
    .filter((item) => item?.sheetIndex === state.activeSheetIndex && item.id)
    .map((item) => item.id);
  if (!ids.length && state.selectedImage?.id) ids.push(state.selectedImage.id);
  return [...new Set(ids)];
}



function selectedInkStrokeIdSet() {
  return new Set((state.selectedInkStrokeIds || []).map(String));
}



function selectedInkStrokes(sheet = activeSheet()) {
  const selectedIds = selectedInkStrokeIdSet();
  if (!selectedIds.size) return [];
  return normalizeInkStrokes(sheet?.inkStrokes).filter((stroke) => selectedIds.has(String(stroke.id)));
}



function selectedImageObject() {
  if (!state.selectedImage || state.selectedImage.sheetIndex !== state.activeSheetIndex) return null;
  return activeSheet().images?.find((image) => image.id === state.selectedImage.id) || null;
}



function selectedImageObjects() {
  const ids = activeSelectedImageIds();
  if (!ids.length) return [];
  const imageById = new Map((activeSheet().images || []).map((image) => [image.id, image]));
  return ids.map((id) => imageById.get(id)).filter(Boolean);
}



function selectedImageObjectById(imageId) {
  if (!imageId) return null;
  return activeSheet().images?.find((image) => image.id === imageId) || null;
}



function setSelectedImageIds(ids, primaryId = null) {
  const sheet = activeSheet();
  const validIds = [...new Set(ids || [])].filter((id) => sheet.images?.some((image) => image.id === id));
  const primary = primaryId && validIds.includes(primaryId) ? primaryId : validIds[0] || null;
  state.selectedImage = primary ? { sheetIndex: state.activeSheetIndex, id: primary } : null;
  state.selectedImages = validIds.map((id) => ({ sheetIndex: state.activeSheetIndex, id }));
  state.shapePointEditIds = (state.shapePointEditIds || []).filter((id) => validIds.includes(id));
  if (!primary || state.pictureCrop?.sheetIndex !== state.activeSheetIndex || state.pictureCrop?.imageId !== primary) {
    clearPictureCropState();
  }
  if (validIds.length) state.selectedInkStrokeIds = [];
}



function imageBoxToPixels(sheet, image) {
  if (!image?.tl) return null;
  let left = coordinateToPixels(sheet, "col", image.tl.col, image.tl.nativeColOff, image.tl.nativeCol);
  const top = coordinateToPixels(sheet, "row", image.tl.row, image.tl.nativeRowOff, image.tl.nativeRow);
  let width = Number(image.ext?.width || 0);
  let height = Number(image.ext?.height || 0);
  if ((!width || !height) && image.br) {
    const right = coordinateToPixels(sheet, "col", image.br.col, image.br.nativeColOff, image.br.nativeCol);
    width = Math.abs(right - left);
    left = Math.min(left, right);
    height = coordinateToPixels(sheet, "row", image.br.row, image.br.nativeRowOff, image.br.nativeRow) - top;
  }
  if (width <= 0 || height <= 0) return null;
  return { left, top, width, height };
}



function coordinateToPixels(sheet, axis, value, nativeOffset = 0, nativeIndex = null) {
  const isColumn = axis === "col";
  if (isColumn && sheetColumnsRightToLeft(sheet)) {
    return rtlColumnCoordinateToPixels(sheet, value, nativeOffset, nativeIndex);
  }
  const headerOffset = isColumn ? rowHeaderWidthForSheet(sheet) : COLUMN_HEADER_HEIGHT;
  const items = isColumn ? sheet.columns : sheet.rows;
  const defaultSize = isColumn ? sheet.defaultColumnWidth : sheet.defaultRowHeight;
  const hasNativeAnchor = nativeIndex !== null && Number.isFinite(Number(nativeIndex));
  const coordinate = hasNativeAnchor ? Number(nativeIndex) : Number.isFinite(Number(value)) ? Number(value) : 0;
  const whole = Math.max(0, Math.floor(coordinate));
  const fraction = hasNativeAnchor ? 0 : Math.max(0, coordinate - whole);
  let pixels = headerOffset;

  for (let index = 0; index < whole; index += 1) {
    const item = items[index];
    pixels += item?.hidden ? 0 : item?.width || item?.height || defaultSize || 0;
  }

  const current = items[whole];
  const currentSize = current?.hidden ? 0 : current?.width || current?.height || defaultSize || 0;
  pixels += fraction * currentSize;
  if (hasNativeAnchor && nativeOffset) pixels += (Number(nativeOffset) / EMU_PER_PIXEL) * (sheet.displayScale || 1);
  return Math.round(pixels * 100) / 100;
}



function sheetColumnsRightToLeft(sheet) {
  return sheet?.rightToLeft === true;
}



function visualGridColumnStart(sheet, col, colSpan = 1) {
  const safeCol = Math.max(1, Math.round(Number(col) || 1));
  const safeSpan = Math.max(1, Math.round(Number(colSpan) || 1));
  if (!sheetColumnsRightToLeft(sheet)) return safeCol + 1;
  const rightCol = Math.min(Math.max(1, Number(sheet?.colCount || safeCol)), safeCol + safeSpan - 1);
  return Math.max(2, Number(sheet?.colCount || rightCol) - rightCol + 2);
}



function rtlColumnCoordinateToPixels(sheet, value, nativeOffset = 0, nativeIndex = null) {
  const headerOffset = sheetHeadingsVisible(sheet) ? rowHeaderWidthForSheet(sheet) : 0;
  const colCount = Math.max(1, Number(sheet?.colCount || sheet?.columns?.length || 1));
  const hasNativeAnchor = nativeIndex !== null && Number.isFinite(Number(nativeIndex));
  if (hasNativeAnchor) {
    const nativeCol = Math.max(1, Math.min(colCount, Math.floor(Number(nativeIndex)) + 1));
    const offsetPx = nativeOffset ? (Number(nativeOffset) / EMU_PER_PIXEL) * (sheet.displayScale || 1) : 0;
    return Math.round((headerOffset + widthBetweenColumns(sheet, nativeCol + 1, colCount) + offsetPx) * 100) / 100;
  }

  const coordinate = Number.isFinite(Number(value)) ? Number(value) : 0;
  const whole = Math.max(0, Math.min(colCount, Math.floor(coordinate)));
  const fraction = Math.max(0, Math.min(1, coordinate - whole));
  const boundary = (count) => {
    if (count <= 0) return headerOffset + widthBetweenColumns(sheet, 1, colCount);
    if (count >= colCount) return headerOffset;
    return headerOffset + widthBetweenColumns(sheet, count + 1, colCount);
  };
  const start = boundary(whole);
  const end = boundary(Math.min(colCount, whole + 1));
  return Math.round((start + (end - start) * fraction) * 100) / 100;
}



function cellInnerHtml(cell, display, overflowWidth = null, sheet = activeSheet()) {
  const text = display == null ? "" : String(display);
  const indicators = cellAnnotationMarkersHtml(cell);
  const filterButton = cellAutoFilterButtonHtml(cell);
  if (cellShowsValidationDropdown(cell, sheet)) {
    return `${indicators}${filterButton}<span class="cell-content" ${cellContentStyle(overflowWidth)}>${escapeHtml(text)}</span>${cellValidationDropdownButtonHtml(cell)}`;
  }

  if (cell.hyperlink) {
    const hyperlink = escapeAttr(cell.hyperlink);
    return `${indicators}${filterButton}<a class="cell-link cell-content" ${cellContentStyle(overflowWidth)} href="${hyperlink}" data-hyperlink="${hyperlink}" target="_blank" rel="noopener noreferrer">${escapeHtml(text)}</a>`;
  }

  return `${indicators}${filterButton}<span class="cell-content" ${cellContentStyle(overflowWidth)}>${escapeHtml(text)}</span>`;
}



function cellValidationDropdownButtonHtml(cell) {
  const label = `${columnName(cell?.col || 1)}${cell?.row || 1} ドロップダウン リストを表示`;
  return `<button class="cell-validation-dropdown-button" type="button" aria-label="${escapeAttr(label)}" title="ドロップダウン リストを表示"></button>`;
}



function cellAutoFilterButtonHtml(cell) {
  if (!cell || !isAutoFilterHeaderCell(activeSheet(), cell.row, cell.col)) return "";
  const criterion = normalizeAutoFilterCriterion(autoFilterStateForHeaderCell(activeSheet(), cell.row, cell.col)?.criteria?.[cell.col] || null);
  const filtered = criterion.kind !== "all";
  const sortKey = autoFilterSortKeyForHeaderCell(activeSheet(), cell.row, cell.col);
  const sortDirection = sortKey?.direction === "desc" ? "desc" : sortKey ? "asc" : "";
  const sortClass = sortDirection ? ` is-sorted is-sort-${sortDirection}` : "";
  const labelParts = ["フィルター"];
  if (filtered) labelParts.push("適用済み");
  if (sortDirection) labelParts.push(sortDirection === "desc" ? "降順で並べ替え済み" : "昇順で並べ替え済み");
  return `<button class="cell-filter-button${filtered ? " is-filtered" : ""}${sortClass}" type="button" data-filter-row="${cell.row}" data-filter-col="${cell.col}"${sortDirection ? ` data-sort-direction="${sortDirection}"` : ""} aria-label="${escapeAttr(labelParts.join("、"))}" title="${escapeAttr(labelParts.join("、"))}"></button>`;
}



function cellAnnotationTitle(cell) {
  const lines = [];
  const comment = normalizeCommentText(cell?.comment);
  const note = normalizeNote(cell?.note);
  if (comment) lines.push(`コメント:\n${comment}`);
  if (note) lines.push(`メモ:\n${note}`);
  return lines.join("\n\n");
}



function cellHasComment(cell) {
  return Boolean(normalizeCommentText(cell?.comment));
}



function cellHasNote(cell) {
  return Boolean(normalizeNote(cell?.note));
}



function cellHasConstant(cell) {
  if (!cell || cell.formula || cell.dataTableFormula || (typeof cell.raw === "string" && cell.raw.startsWith("="))) return false;
  const value = cell.raw ?? cell.cached ?? cell.display;
  return value != null && String(value).trim() !== "";
}



function cellHasFormula(cell) {
  return Boolean(cell?.formula || cell?.dataTableFormula || (typeof cell?.raw === "string" && cell.raw.trim().startsWith("=")));
}



function cellHasConditionalFormat(cell) {
  return Boolean(cell?.conditionalFormatting || cell?.conditionalFormat || cell?.conditionalFormats || cell?.cfRule || cell?.cfRules);
}



function cellAnnotationMarkersHtml(cell) {
  const markers = [];
  if (cellHasComment(cell)) markers.push('<span class="cell-annotation-marker is-comment" aria-hidden="true"></span>');
  if (cellHasNote(cell)) markers.push('<span class="cell-annotation-marker is-note" aria-hidden="true"></span>');
  return markers.join("");
}



function cellContentStyle(overflowWidth) {
  return overflowWidth ? `style="width:${overflowWidth}px"` : "";
}



function selectCell(cellElement, options = {}) {
  const point = cellPointFromElement(cellElement);
  if (!point) return;
  selectCellPoint(point, options);
}



function selectCellPoint(point, options = {}) {
  const range = selectionRangeFromPoints(point, point);
  if (!ensureSelectableRangeForProtection(activeSheet(), range)) return;
  const selected = setSelectedCell(point.row, point.col, { preserveRange: options.preserveRange });

  updateSelectionUi();
  updateFormulaBarForCell(selected.row, selected.col);
}



function handleSheetHostCellClick(event) {
  if (state.suppressNextCellClick) {
    state.suppressNextCellClick = false;
    return;
  }
  if (state.formatPainter) {
    const formatPoint = formatPainterPointFromPointerEvent(event);
    if (!formatPoint) return;
    event.preventDefault();
    applyFormatPainterAtPoint(formatPoint);
    return;
  }
  const hyperlink = hyperlinkFromCellClickEvent(event);
  if (hyperlink && openHyperlinkTarget(hyperlink)) {
    event.preventDefault();
    event.stopPropagation();
    return;
  }
  const point = cellPointFromPointerEvent(event);
  if (!point) return;
  selectCellPoint(point);
}



function startColumnHeaderSelection(event) {
  if (!state.model || event.button !== 0 || event.target.closest?.(".column-resize-handle")) return;
  const col = Number(event.currentTarget.dataset.col);
  if (!col) return;
  if (startFormulaHeaderReferenceSelection("col", col, event)) return;
  event.preventDefault();
  event.stopPropagation();
  startHeaderSelection("col", col, event);
}



function startRowHeaderSelection(event) {
  if (!state.model || event.button !== 0 || event.target.closest?.(".row-resize-handle")) return;
  const row = Number(event.currentTarget.dataset.row);
  if (!row) return;
  if (startFormulaHeaderReferenceSelection("row", row, event)) return;
  event.preventDefault();
  event.stopPropagation();
  startHeaderSelection("row", row, event);
}



function startHeaderSelection(kind, index, event) {
  const additive = event.ctrlKey || event.metaKey;
  const anchor = headerSelectionAnchor(kind, index, event.shiftKey);
  const baseRanges = additive ? rawActiveSelectionRanges() : [];
  const range = headerSelectionRange(kind, anchor, index);
  if (!ensureSelectableRangeForProtection(activeSheet(), range)) return;
  const point = kind === "col" ? { sheetIndex: state.activeSheetIndex, row: 1, col: index } : { sheetIndex: state.activeSheetIndex, row: index, col: 1 };

  applySelectionRange(range, point, { additive, baseRanges });
  state.selectionAnchor = kind === "col"
    ? { sheetIndex: state.activeSheetIndex, row: 1, col: anchor }
    : { sheetIndex: state.activeSheetIndex, row: anchor, col: 1 };
  state.selectionStart = state.selectionAnchor;
  state.selectionDrag = { kind, additive, baseRanges, anchor };
  state.isSelecting = true;
  updateSelectionUi();
  updateFormulaBarForSelection();
}



function extendHeaderSelection(event) {
  if (extendFormulaHeaderReferenceSelection(event)) return;
  const drag = state.selectionDrag;
  if (!state.isSelecting || !drag || (drag.kind !== "col" && drag.kind !== "row")) return;
  const index = Number(drag.kind === "col" ? event.currentTarget.dataset.col : event.currentTarget.dataset.row);
  if (!index) return;
  applyHeaderSelectionIndex(drag, index);
}



function updateHeaderSelectionDrag(event) {
  const drag = state.selectionDrag;
  if (!state.isSelecting || !drag || (drag.kind !== "col" && drag.kind !== "row")) return;
  if (event.buttons !== undefined && (event.buttons & 1) !== 1) return;
  event.preventDefault?.();
  updateHeaderSelectionFromClient(event.clientX, event.clientY, { clampToGrid: true });
  updateSelectionAutoScroll(event.clientX, event.clientY);
}



function updateHeaderSelectionFromClient(clientX, clientY, options = {}) {
  const drag = state.selectionDrag;
  if (!state.isSelecting || !drag || (drag.kind !== "col" && drag.kind !== "row")) return false;
  const sheetPixel = sheetPixelFromClientPosition(clientX, clientY, options);
  if (!sheetPixel) return false;
  const index = sheetIndexFromPixel(activeSheet(), drag.kind, drag.kind === "col" ? sheetPixel.x : sheetPixel.y);
  if (!index) return false;
  return applyHeaderSelectionIndex(drag, index);
}



function applyHeaderSelectionIndex(drag, index) {
  const range = headerSelectionRange(drag.kind, drag.anchor, index);
  if (!ensureSelectableRangeForProtection(activeSheet(), range)) return false;
  const point = drag.kind === "col"
    ? { sheetIndex: state.activeSheetIndex, row: 1, col: index }
    : { sheetIndex: state.activeSheetIndex, row: index, col: 1 };
  applySelectionRange(range, point, { additive: drag.additive, baseRanges: drag.baseRanges });
  state.suppressNextCellClick = true;
  updateSelectionUi();
  updateFormulaBarForSelection();
  return true;
}



function headerSelectionAnchor(kind, index, useShiftRange) {
  if (!useShiftRange) return index;
  const range = activeSelectionRange();
  if (kind === "col") {
    if (range?.fullColumn) return state.selectionStart?.col || range.left;
    return state.selected?.sheetIndex === state.activeSheetIndex ? state.selected.col : index;
  }
  if (range?.fullRow) return state.selectionStart?.row || range.top;
  return state.selected?.sheetIndex === state.activeSheetIndex ? state.selected.row : index;
}



function headerSelectionRange(kind, start, end) {
  return kind === "col" ? columnSelectionRange(start, end) : rowSelectionRange(start, end);
}



function columnSelectionRange(startCol, endCol) {
  const sheet = activeSheet();
  return {
    sheetIndex: state.activeSheetIndex,
    top: 1,
    left: Math.min(startCol, endCol),
    bottom: sheet.rowCount,
    right: Math.max(startCol, endCol),
    fullColumn: true,
  };
}



function rowSelectionRange(startRow, endRow) {
  const sheet = activeSheet();
  return {
    sheetIndex: state.activeSheetIndex,
    top: Math.min(startRow, endRow),
    left: 1,
    bottom: Math.max(startRow, endRow),
    right: sheet.colCount,
    fullRow: true,
  };
}



function setSelectedCell(row, col, options = {}) {
  const point = selectableCellPoint(row, col);
  state.selected = point;
  state.selectedImage = null;
  state.selectedImages = [];
  state.formulaEdit = null;
  stopFormulaReferenceAutoScroll();
  state.formulaReferenceRange = null;
  state.formulaReferenceDrag = null;
  state.formulaReferenceSheetNavigation = null;
  if (!options.preserveRange) {
    state.selectionRange = selectionRangeFromPoints(point, point);
    state.selectionRanges = [state.selectionRange];
    state.selectionAnchor = point;
    state.selectionStart = point;
  }
  return point;
}



function applySelectionRange(range, point, options = {}) {
  state.selected = point;
  state.selectedImage = null;
  state.selectedImages = [];
  state.formulaEdit = null;
  stopFormulaReferenceAutoScroll();
  state.formulaReferenceRange = null;
  state.formulaReferenceDrag = null;
  state.formulaReferenceSheetNavigation = null;
  state.selectionRange = range;
  state.selectionRanges = options.additive ? [...(options.baseRanges || rawActiveSelectionRanges()), range] : [range];
}



function rawActiveSelectionRanges() {
  if (state.selectionRange?.sheetIndex !== state.activeSheetIndex) return [];
  const ranges = (state.selectionRanges || []).filter((range) => range?.sheetIndex === state.activeSheetIndex);
  if (!ranges.length || !ranges.some((range) => rawRangesEqual(range, state.selectionRange))) {
    return state.selectionRange ? [state.selectionRange] : [];
  }
  return ranges;
}



function moveSelectionByKeyboard(direction) {
  if (!state.selected || state.selected.sheetIndex !== state.activeSheetIndex) return;

  const sheet = activeSheet();
  const currentRange = activeSelectionRange() || selectionRangeFromPoints(state.selected, state.selected);
  if (!currentRange) return;

  let targetRow = state.selected.row;
  let targetCol = state.selected.col;
  switch (direction) {
    case "up":
      targetRow = Math.max(1, currentRange.top - 1);
      targetCol = Math.min(Math.max(1, targetCol), sheet.colCount);
      break;
    case "down":
      targetRow = currentRange.bottom + 1;
      targetCol = Math.min(Math.max(1, targetCol), sheet.colCount);
      break;
    case "left":
      targetRow = Math.min(Math.max(1, targetRow), sheet.rowCount);
      targetCol = Math.max(1, currentRange.left - 1);
      break;
    case "right":
      targetRow = Math.min(Math.max(1, targetRow), sheet.rowCount);
      targetCol = currentRange.right + 1;
      break;
    default:
      return;
  }

  selectCellByKeyboard(targetRow, targetCol, {
    rowStep: direction === "up" ? -1 : direction === "down" ? 1 : 0,
    colStep: direction === "left" ? -1 : direction === "right" ? 1 : 0,
  });
}



function keyboardNavigationTargetFromEvent(event) {
  if (!state.selected || state.selected.sheetIndex !== state.activeSheetIndex) return null;
  if (event.altKey) return null;

  const key = String(event.key || "");
  const sheet = activeSheet();
  const currentRange = activeSelectionRange() || selectionRangeFromPoints(state.selected, state.selected);
  if (!currentRange) return null;
  const isModifier = event.ctrlKey || event.metaKey;

  if (!isModifier && key === "Tab") {
    return {
      row: state.selected.row,
      col: event.shiftKey ? Math.max(1, currentRange.left - 1) : currentRange.right + 1,
      colStep: event.shiftKey ? -1 : 1,
    };
  }

  if (!event.shiftKey && key === "Home") {
    return isModifier ? { row: 1, col: 1, rowStep: 1, colStep: 1 } : { row: state.selected.row, col: 1, colStep: 1 };
  }

  if (!event.shiftKey && key === "End") {
    if (isModifier) {
      const used = keyboardNavigationUsedRange(sheet);
      return { row: used.bottom, col: used.right, rowStep: -1, colStep: -1 };
    }
    return { row: state.selected.row, col: keyboardNavigationLastUsedColumnInRow(sheet, state.selected.row, state.selected.col), colStep: -1 };
  }

  if (!isModifier && !event.shiftKey && (key === "PageDown" || key === "PageUp")) {
    const step = keyboardNavigationPageRowStep(sheet);
    return {
      row: key === "PageDown" ? currentRange.bottom + step : Math.max(1, currentRange.top - step),
      col: state.selected.col,
      rowStep: key === "PageDown" ? 1 : -1,
    };
  }

  if (isModifier && !event.shiftKey) {
    const direction = {
      ArrowDown: "down",
      ArrowUp: "up",
      ArrowLeft: "left",
      ArrowRight: "right",
    }[key];
    if (direction) return keyboardNavigationEdgeTarget(sheet, currentRange, direction);
  }

  return null;
}



function keyboardSelectionExtensionTargetFromEvent(event) {
  if (!state.selected || state.selected.sheetIndex !== state.activeSheetIndex) return null;
  if (!event.shiftKey || event.altKey || isEnterKeyEvent(event) || String(event.key || "") === "Tab") return null;

  const key = String(event.key || "");
  const sheet = activeSheet();
  const selectedRange = selectionRangeFromPoints(state.selected, state.selected);
  const currentRange = activeSelectionRange() || selectedRange;
  const isModifier = event.ctrlKey || event.metaKey;

  if (isModifier) {
    if (key === "Home") return { row: 1, col: 1, rowStep: 1, colStep: 1 };
    if (key === "End") {
      const used = keyboardNavigationUsedRange(sheet);
      return { row: used.bottom, col: used.right, rowStep: -1, colStep: -1 };
    }
    const direction = {
      ArrowDown: "down",
      ArrowUp: "up",
      ArrowLeft: "left",
      ArrowRight: "right",
    }[key];
    if (direction) return keyboardNavigationEdgeTarget(sheet, currentRange, direction);
    return null;
  }

  if (key === "Home") return { row: state.selected.row, col: 1, colStep: 1 };
  if (key === "End") {
    return { row: state.selected.row, col: keyboardNavigationLastUsedColumnInRow(sheet, state.selected.row, state.selected.col), colStep: -1 };
  }
  if (key === "PageDown" || key === "PageUp") {
    const step = keyboardNavigationPageRowStep(sheet);
    return {
      row: key === "PageDown" ? selectedRange.bottom + step : Math.max(1, selectedRange.top - step),
      col: state.selected.col,
      rowStep: key === "PageDown" ? 1 : -1,
    };
  }
  const direction = {
    ArrowDown: "down",
    ArrowUp: "up",
    ArrowLeft: "left",
    ArrowRight: "right",
  }[key];
  if (!direction) return null;
  return keyboardSelectionAdjacentTarget(selectedRange, direction);
}



function keyboardSelectionAdjacentTarget(selectedRange, direction) {
  switch (direction) {
    case "up":
      return { row: Math.max(1, selectedRange.top - 1), col: state.selected.col, rowStep: -1 };
    case "down":
      return { row: selectedRange.bottom + 1, col: state.selected.col, rowStep: 1 };
    case "left":
      return { row: state.selected.row, col: Math.max(1, selectedRange.left - 1), colStep: -1 };
    case "right":
      return { row: state.selected.row, col: selectedRange.right + 1, colStep: 1 };
    default:
      return null;
  }
}



function keyboardNavigationPageRowStep(sheet) {
  const visible = visibleGridRange(sheet);
  const visibleRows = Math.max(1, visible.bottom - visible.top + 1 - VIRTUAL_ROW_OVERSCAN * 2);
  return Math.max(1, visibleRows - 1);
}



function keyboardNavigationUsedRange(sheet) {
  const used = { bottom: 1, right: 1 };
  Object.values(sheet?.cells || {}).forEach((cell) => {
    if (!cell) return;
    const row = Number(cell.row) || 1;
    const col = Number(cell.col) || 1;
    used.bottom = Math.max(used.bottom, row);
    used.right = Math.max(used.right, col);
  });
  (sheet?.merges || []).forEach((merge) => {
    used.bottom = Math.max(used.bottom, merge.bottom);
    used.right = Math.max(used.right, merge.right);
  });
  (sheet?.tables || []).forEach((table) => {
    if (!table?.range) return;
    used.bottom = Math.max(used.bottom, table.range.bottom);
    used.right = Math.max(used.right, table.range.right);
  });
  (sheet?.autoFilters || []).forEach((filter) => {
    if (!filter?.range) return;
    used.bottom = Math.max(used.bottom, filter.range.bottom);
    used.right = Math.max(used.right, filter.range.right);
  });
  return used;
}



function keyboardNavigationLastUsedColumnInRow(sheet, row, fallbackCol) {
  let last = 0;
  Object.values(sheet?.cells || {}).forEach((cell) => {
    if (!cell || Number(cell.row) !== Number(row)) return;
    if (!cellHasKeyboardNavigationContent(cell)) return;
    last = Math.max(last, Number(cell.col) || 0);
  });
  (sheet?.merges || []).forEach((merge) => {
    if (row < merge.top || row > merge.bottom) return;
    const master = sheet.cells?.[cellKey(merge.top, merge.left)];
    if (!cellHasKeyboardNavigationContent(master)) return;
    last = Math.max(last, merge.right);
  });
  return Math.max(1, last || fallbackCol || 1);
}



function keyboardNavigationEdgeTarget(sheet, currentRange, direction) {
  const used = keyboardNavigationUsedRange(sheet);
  const row = state.selected.row;
  const col = state.selected.col;
  if (direction === "right") {
    return { row, col: keyboardNavigationFindEdgeInRow(sheet, row, currentRange.right, 1, Math.max(sheet.colCount, used.right, currentRange.right)), colStep: 1 };
  }
  if (direction === "left") {
    return { row, col: keyboardNavigationFindEdgeInRow(sheet, row, currentRange.left, -1, Math.max(sheet.colCount, used.right, currentRange.left)), colStep: -1 };
  }
  if (direction === "down") {
    return { row: keyboardNavigationFindEdgeInColumn(sheet, col, currentRange.bottom, 1, Math.max(sheet.rowCount, used.bottom, currentRange.bottom)), col, rowStep: 1 };
  }
  if (direction === "up") {
    return { row: keyboardNavigationFindEdgeInColumn(sheet, col, currentRange.top, -1, Math.max(sheet.rowCount, used.bottom, currentRange.top)), col, rowStep: -1 };
  }
  return null;
}



function keyboardNavigationFindEdgeInRow(sheet, row, startCol, step, maxCol) {
  const minCol = 1;
  const hasStartContent = keyboardNavigationCellHasContent(sheet, row, startCol);
  if (hasStartContent) {
    let col = startCol;
    while (col + step >= minCol && col + step <= maxCol && keyboardNavigationCellHasContent(sheet, row, col + step)) {
      col += step;
    }
    return Math.max(minCol, Math.min(maxCol, col));
  }

  let col = startCol + step;
  while (col >= minCol && col <= maxCol) {
    if (keyboardNavigationCellHasContent(sheet, row, col)) return col;
    col += step;
  }
  return step > 0 ? maxCol : minCol;
}



function keyboardNavigationFindEdgeInColumn(sheet, col, startRow, step, maxRow) {
  const minRow = 1;
  const hasStartContent = keyboardNavigationCellHasContent(sheet, startRow, col);
  if (hasStartContent) {
    let row = startRow;
    while (row + step >= minRow && row + step <= maxRow && keyboardNavigationCellHasContent(sheet, row + step, col)) {
      row += step;
    }
    return Math.max(minRow, Math.min(maxRow, row));
  }

  let row = startRow + step;
  while (row >= minRow && row <= maxRow) {
    if (keyboardNavigationCellHasContent(sheet, row, col)) return row;
    row += step;
  }
  return step > 0 ? maxRow : minRow;
}



function keyboardNavigationCellHasContent(sheet, row, col) {
  if (isSheetAxisIndexHidden(sheet, "row", row) || isSheetAxisIndexHidden(sheet, "col", col)) return false;
  const merge = findMergeForCell(sheet, row, col);
  const sourceRow = merge?.top || row;
  const sourceCol = merge?.left || col;
  return cellHasKeyboardNavigationContent(sheet?.cells?.[cellKey(sourceRow, sourceCol)]);
}



function cellHasKeyboardNavigationContent(cell) {
  if (!cell) return false;
  if (cell.formula || cell.dataTableFormula) return true;
  const value = cell.raw ?? cell.cached ?? cell.display;
  if (value == null) return false;
  return String(value).trim() !== "";
}



function extendSelectionByKeyboard(row, col, options = {}) {
  if (!state.selected || state.selected.sheetIndex !== state.activeSheetIndex) return;
  const sheet = activeSheet();
  const minRows = Math.max(1, row);
  const minCols = Math.max(1, col);
  const willExtend = minRows > sheet.rowCount || minCols > sheet.colCount;
  ensureSheetSize(sheet, minRows, minCols);

  const anchor = state.selectionStart?.sheetIndex === state.activeSheetIndex
    ? state.selectionStart
    : { ...state.selected };
  const target = visibleCellTarget(sheet, row, col, options);
  const point = selectableCellPoint(target.row, target.col);
  const range = selectionRangeFromPoints(anchor, point);
  if (!ensureSelectableRangeForProtection(sheet, range)) return;

  state.selected = point;
  state.selectedImage = null;
  state.selectedImages = [];
  state.formulaEdit = null;
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionAnchor = anchor;
  state.selectionStart = anchor;

  if (willExtend) renderSheet();
  const didScroll = scrollRangeIntoView(range);
  const needsRender = willExtend || didScroll || !rangeIntersects(range, visibleGridRange(sheet));
  if (needsRender) renderSheet();
  updateSelectionUi();
  updateFormulaBarForCell(point.row, point.col);
}



function selectCellByKeyboard(row, col, options = {}) {
  const sheet = activeSheet();
  const minRows = Math.max(1, row);
  const minCols = Math.max(1, col);
  const willExtend = minRows > sheet.rowCount || minCols > sheet.colCount;
  ensureSheetSize(sheet, minRows, minCols);

  const target = visibleCellTarget(sheet, row, col, options);
  const point = setSelectedCell(target.row, target.col);
  const range = activeSelectionRange();
  if (willExtend) renderSheet();
  const didScroll = scrollRangeIntoView(range);
  const needsRender = willExtend || didScroll || !rangeIntersects(range, visibleGridRange(sheet));

  if (needsRender) {
    renderSheet();
  }
  updateSelectionUi();
  updateFormulaBarForCell(point.row, point.col);
}



function visibleCellTarget(sheet, row, col, options = {}) {
  const rowStep = Number(options.rowStep) || 0;
  const colStep = Number(options.colStep) || 0;
  return {
    row: visibleSheetAxisIndex(sheet, "row", Math.max(1, Number(row) || 1), rowStep),
    col: visibleSheetAxisIndex(sheet, "col", Math.max(1, Number(col) || 1), colStep),
  };
}



function visibleSheetAxisIndex(sheet, axis, index, preferredStep = 0) {
  const max = Math.max(1, axis === "row" ? sheet.rowCount : sheet.colCount, index);
  const start = Math.max(1, Math.min(max, Number(index) || 1));
  if (!isSheetAxisIndexHidden(sheet, axis, start)) return start;
  const primaryStep = preferredStep < 0 ? -1 : 1;
  const primary = firstVisibleSheetAxisIndex(sheet, axis, start, primaryStep, max);
  if (primary) return primary;
  return firstVisibleSheetAxisIndex(sheet, axis, start, -primaryStep, max) || start;
}



function firstVisibleSheetAxisIndex(sheet, axis, start, step, max) {
  for (let index = start; index >= 1 && index <= max; index += step) {
    if (!isSheetAxisIndexHidden(sheet, axis, index)) return index;
  }
  return null;
}



function isSheetAxisIndexHidden(sheet, axis, index) {
  const item = axis === "row" ? sheet?.rows?.[index - 1] : sheet?.columns?.[index - 1];
  return Boolean(item?.hidden);
}



function scrollRangeIntoView(range) {
  const host = $sheetHost[0];
  if (!host || !range) return false;

  const sheet = activeSheet();
  const left = coordinateToPixels(sheet, "col", range.left - 1);
  const top = coordinateToPixels(sheet, "row", range.top - 1);
  const right = coordinateToPixels(sheet, "col", range.right);
  const bottom = coordinateToPixels(sheet, "row", range.bottom);
  const margin = 8;
  const previousLeft = host.scrollLeft;
  const previousTop = host.scrollTop;
  const rangeWidth = right - left;
  const rangeHeight = bottom - top;

  if (rangeWidth + margin * 2 >= host.clientWidth) {
    host.scrollLeft = Math.max(0, left - margin);
  } else if (left < host.scrollLeft) {
    host.scrollLeft = Math.max(0, left - margin);
  } else if (right > host.scrollLeft + host.clientWidth) {
    host.scrollLeft = Math.max(0, right - host.clientWidth + margin);
  }

  if (rangeHeight + margin * 2 >= host.clientHeight) {
    host.scrollTop = Math.max(0, top - margin);
  } else if (top < host.scrollTop) {
    host.scrollTop = Math.max(0, top - margin);
  } else if (bottom > host.scrollTop + host.clientHeight) {
    host.scrollTop = Math.max(0, bottom - host.clientHeight + margin);
  }

  const changed = host.scrollLeft !== previousLeft || host.scrollTop !== previousTop;
  if (changed) forceSyncHorizontalScrollbarFromHost();
  return changed;
}



function updateFormulaBarForSelection() {
  const point = selectionToolbarPoint();
  if (!point) return;
  updateFormulaBarForCell(point.row, point.col);
}



function selectionToolbarPoint() {
  const range = activeSelectionRanges()[0] || activeSelectionRange();
  if (!range) return state.selected;
  if (state.selectionStart?.sheetIndex === state.activeSheetIndex && isCellInRange(range, state.selectionStart.row, state.selectionStart.col)) {
    return state.selectionStart;
  }
  if (state.selectionAnchor?.sheetIndex === state.activeSheetIndex && isCellInRange(range, state.selectionAnchor.row, state.selectionAnchor.col)) {
    return state.selectionAnchor;
  }
  return { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
}



function updateFormulaBarForCell(row, col) {
  $formulaInput.prop("disabled", false).val(getCellEditInput(activeSheet(), row, col));
  $formulaInput[0].scrollTop = 0;
  $formulaExpandButton.prop("disabled", false);
  syncRibbonControlsForCell(row, col);
  syncDataRibbonControls();
}



function beginNameBoxEdit(event) {
  if (!state.model) return;
  const element = event.currentTarget;
  if (event.type === "pointerdown") event.preventDefault();
  if (element.dataset.nameBoxEditing === "true") {
    element.focus({ preventScroll: true });
    requestAnimationFrame(() => selectElementText(element));
    return;
  }
  element.dataset.nameBoxEditing = "true";
  element.dataset.beforeEdit = element.textContent.trim();
  element.setAttribute("contenteditable", "true");
  element.classList.add("is-editing");
  element.focus({ preventScroll: true });
  requestAnimationFrame(() => selectElementText(element));
}



function selectElementText(element) {
  const selection = window.getSelection?.();
  if (!selection || !element) return;
  const range = document.createRange();
  range.selectNodeContents(element);
  selection.removeAllRanges();
  selection.addRange(range);
}



function handleNameBoxKeydown(event) {
  if (!state.model) return;
  if (isElementImeCompositionEvent(event)) return;
  if (event.altKey && !event.ctrlKey && !event.metaKey && event.key === "ArrowDown") {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    openNameBoxMenu();
    return;
  }
  if (event.key === "Enter") {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    finishNameBoxEdit(event.currentTarget, { apply: true });
    return;
  }
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    finishNameBoxEdit(event.currentTarget, { apply: false });
  }
}



function openNameBoxMenu() {
  if (!state.model) return false;
  const definitions = nameBoxVisibleDefinitions();
  closeNameBoxMenu({ restoreFocus: false });
  if (!definitions.length) {
    setStatus("使用できる名前がありません。");
    return false;
  }
  const anchor = $selectedCellLabel[0];
  if (!anchor) return false;
  const rect = anchor.getBoundingClientRect();
  const menu = document.createElement("div");
  menu.id = "nameBoxMenu";
  menu.className = "name-box-menu";
  menu.setAttribute("role", "listbox");
  menu.innerHTML = definitions
    .map((definition) => {
      const scope = definition.scope === "sheet"
        ? state.model?.sheets?.[definition.sheetIndex]?.name || "シート"
        : "ブック";
      return `<button type="button" role="option" data-name-box-menu-item="${escapeAttr(definition.name)}">
        <span class="name-box-menu-name">${escapeHtml(definition.name)}</span>
        <span class="name-box-menu-ref">${escapeHtml(scope)} ${escapeHtml(definition.refersTo || "")}</span>
      </button>`;
    })
    .join("");
  document.body.appendChild(menu);
  menu.style.left = `${Math.round(rect.left)}px`;
  menu.style.top = `${Math.round(rect.bottom + 2)}px`;
  menu.style.minWidth = `${Math.max(160, Math.round(rect.width))}px`;
  menu.style.maxHeight = `${Math.max(120, window.innerHeight - rect.bottom - 10)}px`;
  state.nameBoxMenu = menu;
  menu.querySelector("[data-name-box-menu-item]")?.focus({ preventScroll: true });
  setStatus("名前ボックスの一覧を表示しました。");
  return true;
}



function nameBoxVisibleDefinitions() {
  return (state.model?.definedNames || [])
    .filter((definition) => definition && !definition.hidden && definition.name)
    .slice()
    .sort((left, right) => String(left.name).localeCompare(String(right.name), "ja"));
}



function handleNameBoxMenuItemClick(event) {
  const name = event.currentTarget.dataset.nameBoxMenuItem || "";
  if (!name) return;
  event.preventDefault();
  closeNameBoxMenu({ restoreFocus: false });
  $selectedCellLabel.removeAttr("contenteditable").removeClass("is-editing");
  delete $selectedCellLabel[0]?.dataset.nameBoxEditing;
  delete $selectedCellLabel[0]?.dataset.beforeEdit;
  applyNameBoxText(name);
  $sheetHost.trigger("focus");
}



function closeNameBoxMenuOnOutsidePointer(event) {
  if (!state.nameBoxMenu) return;
  if (event.target.closest?.("#nameBoxMenu") || event.target.closest?.("#selectedCellLabel")) return;
  closeNameBoxMenu({ restoreFocus: false });
}



function closeNameBoxMenu(options = {}) {
  const menu = state.nameBoxMenu || document.getElementById("nameBoxMenu");
  if (menu) menu.remove();
  state.nameBoxMenu = null;
  if (options.restoreFocus) $selectedCellLabel.trigger("focus");
}



function finishNameBoxEditOnBlur(event) {
  if (event.currentTarget.dataset.nameBoxEditing !== "true") return;
  finishNameBoxEdit(event.currentTarget, { apply: true, restoreFocus: false });
}



function finishNameBoxEdit(element, options = {}) {
  if (!element || element.dataset.nameBoxFinishing === "true") return;
  element.dataset.nameBoxFinishing = "true";
  const before = element.dataset.beforeEdit || "";
  const text = element.textContent.trim();
  element.removeAttribute("contenteditable");
  element.classList.remove("is-editing");
  delete element.dataset.nameBoxEditing;
  delete element.dataset.beforeEdit;
  if (options.apply === false) {
    element.textContent = selectionLabelText();
    setStatus("名前ボックスの入力をキャンセルしました。");
  } else if (text && text !== before) {
    applyNameBoxText(text);
  } else {
    element.textContent = selectionLabelText();
  }
  delete element.dataset.nameBoxFinishing;
  if (options.restoreFocus !== false) $sheetHost.trigger("focus");
}



function selectionLabelText() {
  const ranges = activeSelectionRanges();
  if (ranges.length) return ranges.map((range) => rangeToLabel(range)).join(",");
  const range = activeSelectionRange();
  if (range) return rangeToLabel(range);
  if (state.selected?.sheetIndex === state.activeSheetIndex) return `${columnName(state.selected.col)}${state.selected.row}`;
  return "A1";
}



function applyNameBoxText(text) {
  const targets = resolveNameBoxTargets(text);
  if (targets.length) {
    activateNameBoxTargets(targets, text);
    return true;
  }
  const validation = validateDefinedName(text);
  if (!validation.ok) {
    $selectedCellLabel.text(selectionLabelText());
    showMessageBox("参照が正しくないか、名前として使用できません。", { title: "Cht WebSheet" });
    return false;
  }
  return defineNameFromNameBox(text);
}



function resolveNameBoxTargets(text) {
  return String(text || "")
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean)
    .map(resolveGoToReference)
    .filter(Boolean);
}



function activateNameBoxTargets(targets, sourceText) {
  const first = targets[0];
  const sheetIndex = first.sheetIndex;
  if (targets.some((target) => target.sheetIndex !== sheetIndex)) {
    showMessageBox("名前ボックスでは、複数シートにまたがる範囲を同時に選択できません。", { title: "Cht WebSheet" });
    $selectedCellLabel.text(selectionLabelText());
    return;
  }
  state.activeSheetIndex = sheetIndex;
  ensureActiveSheetVisible();
  const sheet = activeSheet();
  const ranges = targets.map((target) => {
    ensureSheetSize(sheet, target.range.bottom, target.range.right);
    return { ...target.range, sheetIndex };
  });
  selectRangesForHomeCommand(ranges, `${sourceText} を選択しました。`);
}



function defineNameFromNameBox(name) {
  const ranges = rawActiveSelectionRanges().length ? rawActiveSelectionRanges() : [activeSelectionRange()].filter(Boolean);
  if (!ranges.length) {
    showMessageBox("名前を定義するセル範囲を選択してください。", { title: "Cht WebSheet" });
    $selectedCellLabel.text(selectionLabelText());
    return false;
  }
  const definition = {
    name,
    scope: "workbook",
    sheetIndex: null,
    refersTo: `=${ranges.map((range) => absoluteRangeReference(state.model.sheets[range.sheetIndex], range)).join(",")}`,
    comment: "",
    hidden: false,
  };
  if (definedNameHasConflict(definition)) {
    showMessageBox(definedNameDuplicateMessage(name), { title: "Cht WebSheet" });
    $selectedCellLabel.text(selectionLabelText());
    return false;
  }
  const before = cloneDefinedNames(state.model.definedNames);
  state.model.definedNames = upsertDefinedName(state.model.definedNames, definition);
  state.hf = buildFormulaEngine(state.model);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    definedNamesBefore: before,
    definedNamesAfter: cloneDefinedNames(state.model.definedNames),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges().map((range) => ({ ...range })),
    selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
    label: `${name} 名前の定義`,
  });
  scheduleWorkbookPersistence();
  $selectedCellLabel.text(selectionLabelText());
  setStatus(`${name} を定義しました。`);
  return true;
}



function syncRibbonControlsForCell(row, col) {
  const sheet = activeSheet();
  const cell = sheet.cells[cellKey(row, col)] || {};
  const css = cellEffectiveCss(sheet, cell);
  const fontOption = fontOptionFromCss(css.fontFamily);
  const fontSize = fontSizeOptionFromCss(css.fontSize);

  if (fontOption) {
    $ribbonPanels.find('[data-ribbon-style="fontFamily"]').val(fontOption);
  }
  if (fontSize) {
    $ribbonPanels.find('[data-ribbon-style="fontSize"]').val(fontSize);
  }
  $ribbonPanels.find('[data-ribbon-style="numberFormat"]').val(numberCategoryFromFormat(cell.numFmt));

  setRibbonButtonActive("bold", isBoldCssValue(css.fontWeight));
  setRibbonButtonActive("italic", css.fontStyle === "italic");
  const underlinePreset = underlinePresetForTextDecoration(css.textDecoration);
  setRibbonButtonActive("underline", Boolean(underlinePreset));
  if (underlinePreset) setRibbonUnderlinePreset(underlinePreset.value);
  setRibbonButtonActive("align-left", !css.textAlign || css.textAlign === "left");
  setRibbonButtonActive("align-center", css.textAlign === "center");
  setRibbonButtonActive("align-right", css.textAlign === "right");
  setRibbonButtonActive("valign-top", css.verticalAlign === "top" || css.alignItems === "flex-start");
  setRibbonButtonActive("valign-middle", !css.verticalAlign || css.verticalAlign === "middle" || css.alignItems === "center");
  setRibbonButtonActive("valign-bottom", css.verticalAlign === "bottom" || css.alignItems === "flex-end");
  setRibbonButtonActive("wrap-text", css.whiteSpace === "pre-wrap");
  setRibbonButtonActive("merge-center", Boolean(findMergeForCell(sheet, row, col)));
}



function fontOptionFromCss(fontFamily) {
  const select = $ribbonPanels.find('[data-ribbon-style="fontFamily"]')[0];
  if (!select) return "";
  const normalized = normalizeFontName(fontFamily);
  if (!normalized) return "Calibri";
  const options = [...select.options].map((option) => option.value);
  return (
    options.find((option) => normalized.includes(normalizeFontName(option))) ||
    options.find((option) => normalizeFontName(fontFamilyForExcel(option)).includes(normalized.split(",")[0] || ""))
  );
}



function fontSizeOptionFromCss(fontSize) {
  const select = $ribbonPanels.find('[data-ribbon-style="fontSize"]')[0];
  if (!select) return "";
  const pt = cssFontSizeToPoints(fontSize);
  if (!pt) return "11";
  const rounded = String(Math.round(pt));
  return [...select.options].some((option) => option.value === rounded) ? rounded : "";
}



function cssFontSizeToPoints(fontSize) {
  const value = Number.parseFloat(fontSize);
  if (!Number.isFinite(value) || value <= 0) return null;
  return String(fontSize || "").trim().endsWith("px") ? value * 72 / 96 : value;
}



function normalizeFontName(value) {
  return String(value || "")
    .replaceAll('"', "")
    .replaceAll("'", "")
    .trim()
    .toLowerCase();
}



function isBoldCssValue(value) {
  const numeric = Number.parseInt(value, 10);
  return value === "bold" || value === "700" || (Number.isFinite(numeric) && numeric >= 600);
}



function setRibbonButtonActive(command, active) {
  const button = command === "underline"
    ? $ribbonPanels.find(`[data-ribbon-command="${command}"], [data-ribbon-underline-apply]`)
    : $ribbonPanels.find(`[data-ribbon-command="${command}"]`);
  button.toggleClass("is-active", Boolean(active));
  button.attr("aria-pressed", String(Boolean(active)));
}



function syncDataRibbonControls() {
  const sheet = state.model ? activeSheet() : null;
  const hasWorkbook = Boolean(state.model && sheet);
  const rawSelectionRange = hasWorkbook ? activeSelectionRange() : null;
  const directContext = hasWorkbook ? dataRibbonSelectionContext(sheet) : null;
  const dataContext = directContext;
  const dataRange = dataContext?.range || null;
  const directRange = directContext?.range || null;
  const hasDataRange = Boolean(dataRange && dataRibbonRangeHasData(sheet, dataRange));
  const hasDirectSelection = Boolean(directRange);
  const directSingleColumn = Boolean(directRange && directRange.left === directRange.right && !directRange.fullRow);
  const directHasData = Boolean(directRange && dataRibbonRangeHasData(sheet, directRange));
  const directSingleCell = Boolean(directRange && isSingleCellRange(directRange));
  const shouldResolveFlashFillTarget = Boolean(
    directRange &&
      directSingleColumn &&
      !directSingleCell &&
      !directRange.fullColumn &&
      rangeCellCount(directRange) <= 10000
  );
  const flashFillTargetRange = shouldResolveFlashFillTarget ? flashFillTargetRangeForSelection(sheet, directRange) : directRange;
  const canFlashFill = Boolean(
    flashFillTargetRange &&
      directSingleColumn &&
      flashFillTargetRange.left > 1 &&
      !flashFillTargetRange.fullColumn &&
      (directSingleCell || flashFillRangeHasSourceData(sheet, flashFillTargetRange))
  );
  const forecastable = Boolean(
    dataRange &&
      !isSingleCellRange(dataRange) &&
      !dataRange.fullRow &&
      !dataRange.fullColumn &&
      rangeCellCount(dataRange) <= 10000 &&
      dataRibbonForecastPointCount(sheet, dataRange) >= 2
  );
  const protectedSheet = Boolean(hasWorkbook && isSheetProtected(sheet));
  const protectionAllowsSort = !protectedSheet || sheetProtectionAllows(sheet, "allowSort");
  const protectionAllowsFilter = !protectedSheet || sheetProtectionAllows(sheet, "allowAutoFilter");
  const protectionAllowsFormatCells = !protectedSheet || sheetProtectionAllows(sheet, "allowFormatCells");
  const protectionAllowsInsertRows = !protectedSheet || sheetProtectionAllows(sheet, "allowInsertRows");
  const directRangeEditable = Boolean(directRange && (!protectedSheet || !rangeHasLockedCellsForProtection(sheet, directRange)));
  const dataRangeEditable = Boolean(dataRange && (!protectedSheet || !rangeHasLockedCellsForProtection(sheet, dataRange)));
  const flashFillRangeEditable = Boolean(flashFillTargetRange && (!protectedSheet || !rangeHasLockedCellsForProtection(sheet, flashFillTargetRange)));
  const tableFilters = visibleTableAutoFilters(sheet);
  const hasTableFilterCriteria = tableFilters.some(autoFilterStateHasCriteria);
  const hasTableSortState = Boolean(visibleTableSortStates(sheet).length);
  const hasFilter = Boolean(sheet?.autoFilter?.range || tableFilters.length);
  const hasFilterCriteria = hasTableFilterCriteria || Boolean(
    sheet?.autoFilter?.advancedCriteria ||
    Object.keys(sheet?.autoFilter?.criteria || {}).length,
  );
  const hasSortState = sortStateHasKeys(sheet?.sortState) || hasTableSortState;
  const hasOutline = Boolean(
    sheet?.rows?.some((row) => clampOutlineLevel(row?.outlineLevel)) ||
    sheet?.columns?.some((column) => clampOutlineLevel(column?.outlineLevel)),
  );
  const outlineAxis = dataRange ? outlineAxisForRange(dataRange) : "rows";
  const protectionAllowsOutlineAxis = !protectedSheet || outlineAxisSheetAllows(sheet, outlineAxis);
  const canGroupOutline = Boolean(dataRange && protectionAllowsOutlineAxis);
  const canShowDetail = Boolean(dataRange && protectionAllowsOutlineAxis && outlineDetailTargetFromRange(sheet, dataRange, false));
  const canHideDetail = Boolean(dataRange && protectionAllowsOutlineAxis && outlineDetailTargetFromRange(sheet, dataRange, true));
  const canUngroup = Boolean(dataRange && protectionAllowsOutlineAxis && outlineRangeHasLevel(sheet, dataRange));
  const queryRefs = hasWorkbook ? workbookQueries() : [];
  const hasQueries = Boolean(queryRefs.length);
  const hasRecentSources = Boolean(queryRefs.length || loadRecentDataSources().length);
  const canRefreshAll = queryRefs.some(({ query }) => query?.refreshWithAll !== false);
  const canCreateTableQuery = Boolean(hasDataRange && dataRange && !dataRange.fullRow && !dataRange.fullColumn);
  const canConsolidate = Boolean(hasWorkbook && (!directRange || directRangeEditable));
  const canCreateRelationship = Boolean(hasWorkbook && canCreateWorkbookRelationship());
  const dataRangeIntersectsTable = Boolean(dataRange && rangeIntersectsWorksheetTable(sheet, dataRange));
  const hasEnoughSelectionForDuplicates = Boolean(hasDataRange && dataRange && (dataRange.bottom > dataRange.top || dataRange.right > dataRange.left));
  const hasEnoughSelectionForSubtotal = Boolean(hasDataRange && dataRange && dataRange.bottom > dataRange.top && dataRange.right > dataRange.left);

  setRibbonButtonActive("data-filter", hasFilter);
  setRibbonButtonActive("data-clear-filter", hasFilterCriteria || hasSortState);
  setRibbonButtonActive("data-reapply-filter", hasFilterCriteria || hasSortState);
  setRibbonButtonActive("data-show-detail", canShowDetail);
  setRibbonButtonActive("data-hide-detail", canHideDetail);

  const workbookCommands = [
    "data-get-menu",
    "data-get-text",
    "data-get-web",
    "data-recent-sources",
    "data-existing-connections",
    "data-queries-connections",
    "data-refresh-menu",
    "data-refresh-all",
    "data-sort-asc",
    "data-sort-desc",
    "data-sort-dialog",
    "data-filter",
    "data-advanced-filter",
    "data-validation",
    "data-consolidate",
    "data-relationships",
    "data-manage-data-model",
    "data-what-if",
    "data-analysis",
    "data-group",
  ];
  workbookCommands.forEach((command) => setRibbonCommandDisabled(command, !hasWorkbook));
  setRibbonCommandDisabled("data-get-table-range", !hasWorkbook || !canCreateTableQuery);
  setRibbonCommandDisabled("data-text-to-columns", !hasWorkbook || !hasDirectSelection || !directSingleColumn || !directHasData || !directRangeEditable);
  setRibbonCommandDisabled("data-flash-fill", !hasWorkbook || !hasDirectSelection || !canFlashFill || !flashFillRangeEditable);
  setRibbonCommandDisabled("data-remove-duplicates", !hasWorkbook || !hasEnoughSelectionForDuplicates || !dataRangeEditable);
  setRibbonCommandDisabled("data-consolidate", !hasWorkbook || !canConsolidate);
  setRibbonCommandDisabled("data-relationships", !hasWorkbook || !canCreateRelationship);
  setRibbonCommandDisabled("data-forecast-sheet", !hasWorkbook || !forecastable);
  setRibbonCommandDisabled("data-group", !hasWorkbook || !canGroupOutline);
  setRibbonCommandDisabled("data-ungroup", !hasWorkbook || !canUngroup);
  setRibbonCommandDisabled("data-subtotal", !hasWorkbook || !hasEnoughSelectionForSubtotal || !dataRangeEditable || !protectionAllowsInsertRows || dataRangeIntersectsTable);
  setRibbonCommandDisabled("data-validation", !hasWorkbook || !protectionAllowsFormatCells);
  setRibbonCommandDisabled("data-sort-asc", !hasWorkbook || !protectionAllowsSort);
  setRibbonCommandDisabled("data-sort-desc", !hasWorkbook || !protectionAllowsSort);
  setRibbonCommandDisabled("data-sort-dialog", !hasWorkbook || !protectionAllowsSort);
  setRibbonCommandDisabled("data-filter", !hasWorkbook || !protectionAllowsFilter);
  setRibbonCommandDisabled("data-advanced-filter", !hasWorkbook || !protectionAllowsFilter);
  setRibbonCommandDisabled("data-recent-sources", !hasWorkbook || !hasRecentSources);
  setRibbonCommandDisabled("data-refresh-menu", !hasWorkbook || !hasQueries);
  setRibbonCommandDisabled("data-refresh-all", !hasWorkbook || !canRefreshAll);
  setRibbonCommandDisabled("data-existing-connections", !hasWorkbook || !hasQueries);
  setRibbonCommandDisabled("data-queries-connections", !hasWorkbook || !hasQueries);
  setRibbonCommandDisabled("data-clear-filter", !hasWorkbook || !protectionAllowsFilter || (!hasFilterCriteria && !hasSortState));
  setRibbonCommandDisabled("data-reapply-filter", !hasWorkbook || !protectionAllowsFilter || (!hasFilter && !hasSortState));
  setRibbonCommandDisabled("data-show-detail", !hasWorkbook || !hasOutline || !canShowDetail);
  setRibbonCommandDisabled("data-hide-detail", !hasWorkbook || !hasOutline || !canHideDetail);
  if (!rawSelectionRange) {
    ["data-sort-asc", "data-sort-desc", "data-sort-dialog", "data-filter", "data-advanced-filter", "data-group", "data-ungroup"].forEach((command) => {
      setRibbonCommandDisabled(command, true);
    });
  }
}



function dataRibbonSelectionContext(sheet, options = {}) {
  if (!sheet || !state.model) return null;
  const range = activeSelectionRanges()[0] || activeSelectionRange();
  if (!range) return null;
  let materialized = materializeSelectionRange(range);
  if (!materialized) return null;
  if (options.expandCurrentRegion && isSingleCellRange(materialized)) {
    const point = selectionToolbarPoint() || state.selected || { sheetIndex: state.activeSheetIndex, row: materialized.top, col: materialized.left };
    materialized = currentRegionFromPoint(sheet, point) || materialized;
  }
  return { sheet, range: materialized };
}



function dataRibbonRangeHasData(sheet, range) {
  if (!sheet || !range) return false;
  const cells = sheet.cells || {};
  const cellHasData = (cell) => Boolean(cell?.formula || String(cellDisplayText(cell)).trim());
  const rangeCellTotal = range.fullRow || range.fullColumn ? Infinity : rangeCellCount(range);

  if (Number.isFinite(rangeCellTotal) && rangeCellTotal <= 20000) {
    for (let row = range.top; row <= range.bottom; row += 1) {
      for (let col = range.left; col <= range.right; col += 1) {
        if (cellHasData(cells[cellKey(row, col)])) return true;
      }
    }
    return (sheet.merges || []).some((merge) =>
      rangesIntersect(merge, range) && cellHasData(cells[cellKey(merge.top, merge.left)])
    );
  }

  const used = sheetUsedRange(sheet);
  if (!used) return false;
  if (rangesIntersect({ sheetIndex: state.activeSheetIndex, ...used }, range)) return true;
  if ((sheet.merges || []).some((merge) =>
    rangesIntersect(merge, range) && cellHasData(cells[cellKey(merge.top, merge.left)])
  )) {
    return true;
  }
  return Object.entries(cells).some(([key, cell]) => {
    if (!cellHasData(cell)) return false;
    const [row, col] = parseCellKey(key);
    if (isCellInRange(range, row, col)) return true;
    return (sheet.merges || []).some((merge) =>
      merge.top === row && merge.left === col && rangesIntersect(merge, range)
    );
  });
}



function dataRibbonForecastPointCount(sheet, range) {
  if (!sheet || !range || !dataRibbonRangeHasData(sheet, range)) return 0;
  const sheetId = getSheetId(sheet.name);
  let count = 0;
  for (let row = range.top; row <= range.bottom; row += 1) {
    const yValue = sheet.cells[cellKey(row, range.right)]?.raw ?? getDisplayForCell(sheet, sheetId, row, range.right);
    const y = forecastNumericValue(yValue);
    const xValue = range.right > range.left
      ? sheet.cells[cellKey(row, range.left)]?.raw ?? getDisplayForCell(sheet, sheetId, row, range.left)
      : count + 1;
    const x = range.right > range.left ? forecastTimelineValue(xValue) : Number(xValue);
    if (Number.isFinite(x) && Number.isFinite(y)) count += 1;
    if (count >= 2) return count;
  }
  return count;
}



function outlineRangeHasLevel(sheet, range) {
  if (!sheet || !range) return false;
  const axis = range.fullColumn || (!range.fullRow && range.right - range.left > range.bottom - range.top) ? "columns" : "rows";
  const start = axis === "columns" ? range.left : range.top;
  const end = axis === "columns" ? range.right : range.bottom;
  for (let index = start; index <= end; index += 1) {
    if (outlineLevelAt(sheet, axis, index) > 0) return true;
  }
  return false;
}



function setRibbonCommandDisabled(command, disabled) {
  const button = $ribbonPanels.find(`[data-ribbon-command="${command}"]`);
  button.prop("disabled", Boolean(disabled));
  button.attr("aria-disabled", String(Boolean(disabled)));
}



function startCellEdit(event) {
  const point = cellPointFromPointerEvent(event);
  if (!point) return;
  beginCellEditAtPoint(point, { clientX: event.clientX, clientY: event.clientY });
}



function startSelectedCellEdit(options = {}) {
  if (!state.selected || state.selected.sheetIndex !== state.activeSheetIndex) return null;
  const sheet = activeSheet();
  if (!ensureEditableRangeForProtection(sheet, selectionRangeFromPoints(state.selected, state.selected), "編集")) return null;
  const range = activeSelectionRange() || selectionRangeFromPoints(state.selected, state.selected);
  ensureSheetSize(sheet, range?.bottom || state.selected.row, range?.right || state.selected.col);
  scrollRangeIntoView(range);
  renderSheet();
  updateSelectionUi();
  const cellElement = cellElementForPoint(state.selected);
  if (cellElement) {
    beginCellEdit(cellElement, { caretAtEnd: true, ...options });
    return cellElement;
  }
  return null;
}



function beginCellEditAtPoint(point, options = {}) {
  const sheet = activeSheet();
  if (!ensureEditableRangeForProtection(sheet, selectionRangeFromPoints(point, point), "編集")) return;
  const cellElement = ensureCellElementForPoint(point, { select: true });
  if (cellElement) beginCellEdit(cellElement, options);
}



function ensureCellElementForPoint(point, options = {}) {
  if (!point || point.sheetIndex !== state.activeSheetIndex) return null;
  let cellElement = cellElementForPoint(point);
  if (cellElement) return cellElement;

  const sheet = activeSheet();
  if (!sheet) return null;
  const targetRange = selectionRangeFromPoints(point, point);
  if (options.select) {
    applySelectionRange(targetRange, point);
  }
  if (!rangeIntersects(targetRange, visibleGridRange(sheet))) {
    scrollRangeIntoView(targetRange);
  }
  renderSheet();
  updateSelectionUi();
  cellElement = cellElementForPoint(point);
  return cellElement;
}



function beginCellEdit(cellElement, options = {}) {
  clearAutoFillOptions();
  closeFormulaAutocomplete({ restoreFocus: false });
  clearFormulaReferenceState({ update: false });
  const $cell = $(cellElement);
  const row = Number($cell.data("row"));
  const col = Number($cell.data("col"));
  const sheet = activeSheet();
  if (!ensureEditableRangeForProtection(sheet, selectionRangeFromPoints({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col }), "編集")) return;
  const cell = sheet.cells[cellKey(row, col)] || {};

  cancelViewportRender();
  cancelIdleViewportTrim();
  state.deferredViewportRender = false;
  selectCell(cellElement);
  cellElement.setAttribute("contenteditable", "true");
  applyValidationImeModeToEditable(cellElement, cell.validation);
  const beforeEdit = getCellRawInput(sheet, row, col);
  $cell.data("beforeEdit", beforeEdit);
  $cell.data("beforeRichText", cloneJsonLikeValue(cell.richText || null));
  if (options.replaceText) {
    $cell.removeData("beforeEditText");
    cellElement.textContent = String(options.initialText ?? "");
  } else {
    const beforeEditText = getCellEditInput(sheet, row, col);
    $cell.data("beforeEditText", beforeEditText);
    cellElement.innerHTML = cellEditHtmlForCell(cell, beforeEditText);
  }
  syncEditableCellAlignment(cellElement);
  repaintActiveSheetCanvas();
  cellElement.focus({ preventScroll: true });
  if (options.caretAtEnd) {
    placeCaretAtEnd(cellElement);
  } else {
    placeCaretAtPoint(options.clientX, options.clientY, cellElement);
  }
  saveCellEditSelection(cellElement);
  updateFormulaReferenceUi();
}



function cellEditHtmlForCell(cell, text) {
  const sourceText = String(text ?? "");
  const runs = richTextRunsForDisplay(cell, sourceText);
  if (!runs?.length) return escapeHtml(sourceText);
  return runs.map((run) => cellEditRunHtml(run)).join("");
}



function cellEditRunHtml(run) {
  const text = escapeHtml(run?.text || "");
  const css = normalizeCellEditInlineCss(run?.css || {});
  if (!Object.keys(css).length) return text;
  return `<span style="${escapeAttr(styleObjectToString(css))}">${text}</span>`;
}



function syncEditableCellAlignment(cellElement) {
  if (!cellElement || cellElement.getAttribute("contenteditable") !== "true") return;
  cellElement.style.setProperty("--cell-edit-padding-top", `${CELL_EDIT_MIN_PADDING_TOP}px`);
  const overlay = editableCellOverlayMetrics(cellElement);
  const paddingTop = editableCellPaddingTop(cellElement, overlay);
  const shiftedTextBottom = Number.isFinite(overlay.textBottom)
    ? overlay.textBottom + paddingTop - CELL_EDIT_MIN_PADDING_TOP
    : overlay.height;
  cellElement.style.setProperty("--cell-edit-overlay-left", `${overlay.left}px`);
  cellElement.style.setProperty("--cell-edit-overlay-width", `${overlay.width}px`);
  cellElement.style.removeProperty("--cell-edit-overlay-height");
  cellElement.style.setProperty("--cell-edit-overlay-height", `${Math.ceil(Math.max(overlay.baseHeight || overlay.height, shiftedTextBottom + CELL_EDIT_PADDING_BOTTOM))}px`);
  cellElement.style.setProperty("--cell-edit-padding-top", `${paddingTop}px`);
  cellElement.scrollTop = 0;
}



function editableCellOverlayHeight(cellElement) {
  return editableCellOverlayMetrics(cellElement).height;
}



function editableCellOverlayMetrics(cellElement) {
  const cellRect = cellElement.getBoundingClientRect();
  const baseHeight = Math.max(cellElement.offsetHeight || 0, cellElement.clientHeight || 0, cellRect.height || 0);
  const baseWidth = Math.max(cellElement.offsetWidth || 0, cellElement.clientWidth || 0, cellRect.width || 0);
  const range = document.createRange();
  range.selectNodeContents(cellElement);
  const rects = [...range.getClientRects()].filter((rect) => rect.width > 0 || rect.height > 0);
  range.detach?.();
  if (!rects.length) return { left: 0, width: Math.ceil(baseWidth), height: Math.ceil(baseHeight), baseHeight: Math.ceil(baseHeight) };
  const textLeft = Math.min(...rects.map((rect) => rect.left)) - cellRect.left;
  const textRight = Math.max(...rects.map((rect) => rect.right)) - cellRect.left;
  const textTop = Math.min(...rects.map((rect) => rect.top)) - cellRect.top;
  const textBottom = Math.max(...rects.map((rect) => rect.bottom)) - cellRect.top;
  const left = Math.min(0, Math.floor(textLeft) - 3);
  const right = Math.max(baseWidth, Math.ceil(textRight) + 3);
  return {
    left,
    width: Math.ceil(right - left),
    height: Math.ceil(Math.max(baseHeight, textBottom + 3)),
    baseHeight: Math.ceil(baseHeight),
    textTop,
    textBottom,
  };
}



function editableCellPaddingTop(cellElement, overlay) {
  const baseHeight = Number(overlay?.baseHeight || overlay?.height || 0);
  const textTop = Number(overlay?.textTop);
  const textBottom = Number(overlay?.textBottom);
  const textHeight = Number.isFinite(textTop) && Number.isFinite(textBottom) ? Math.max(0, textBottom - textTop) : 0;
  if (!baseHeight || !textHeight || baseHeight <= textHeight + CELL_EDIT_MIN_PADDING_TOP + CELL_EDIT_PADDING_BOTTOM) {
    return CELL_EDIT_MIN_PADDING_TOP;
  }
  const vertical = editableCellVerticalAlignment(cellElement);
  if (vertical === "bottom") {
    return Math.max(CELL_EDIT_MIN_PADDING_TOP, Math.floor(CELL_EDIT_MIN_PADDING_TOP + baseHeight - textHeight - CELL_EDIT_PADDING_BOTTOM - textTop));
  }
  if (vertical === "middle") {
    return Math.max(CELL_EDIT_MIN_PADDING_TOP, Math.floor(CELL_EDIT_MIN_PADDING_TOP + (baseHeight - textHeight) / 2 - textTop));
  }
  return CELL_EDIT_MIN_PADDING_TOP;
}



function editableCellVerticalAlignment(cellElement) {
  const computed = getComputedStyle(cellElement);
  const alignItems = String(computed.alignItems || cellElement.style.alignItems || "").toLowerCase();
  const verticalAlign = String(computed.verticalAlign || cellElement.style.verticalAlign || "").toLowerCase();
  if (alignItems === "flex-start" || verticalAlign === "top") return "top";
  if (alignItems === "flex-end" || verticalAlign === "bottom") return "bottom";
  return "middle";
}



function applyValidationImeModeToEditable(element, rule = {}) {
  if (!element) return;
  const mode = String(rule?.imeMode || "noControl");
  element.style.imeMode = "";
  element.removeAttribute("inputmode");
  element.removeAttribute("lang");
  if (mode === "disabled") {
    element.style.imeMode = "disabled";
    element.setAttribute("inputmode", "none");
    return;
  }
  if (mode === "off" || mode === "halfAlpha") {
    element.style.imeMode = "inactive";
    element.setAttribute("inputmode", "latin");
    return;
  }
  if (mode === "on" || mode === "hiragana" || mode === "fullKatakana" || mode === "halfKatakana") {
    element.style.imeMode = "active";
    element.setAttribute("lang", "ja");
    element.setAttribute("inputmode", "text");
    return;
  }
  if (mode === "fullAlpha") {
    element.style.imeMode = "active";
    element.setAttribute("inputmode", "text");
    return;
  }
  element.style.imeMode = "auto";
}



function startEditAutoScroll(event) {
  if (event.button !== 0) return;
  state.editAutoScroll = {
    cell: event.currentTarget,
    isCellEdit: event.currentTarget.matches?.(".sheet-cell[contenteditable='true']") === true,
    pointerId: event.pointerId,
    x: event.clientX,
    y: event.clientY,
    outsideCell: false,
  };
  scheduleEditAutoScroll();
}



function updateEditAutoScroll(event) {
  if (!state.editAutoScroll) return;
  if (state.editAutoScroll.pointerId != null && event.pointerId != null && event.pointerId !== state.editAutoScroll.pointerId) return;
  if (event.buttons !== undefined && (event.buttons & 1) !== 1) {
    stopEditAutoScroll();
    return;
  }
  state.editAutoScroll.x = event.clientX;
  state.editAutoScroll.y = event.clientY;
  if (state.editAutoScroll.isCellEdit) {
    state.editAutoScroll.outsideCell = !pointIsInsideElementRect(state.editAutoScroll.cell, event.clientX, event.clientY);
  }
}



function stopEditAutoScroll(event = null) {
  const job = state.editAutoScroll;
  if (
    job?.isCellEdit &&
    event?.type === "pointerup" &&
    (job.pointerId == null || event.pointerId == null || event.pointerId === job.pointerId)
  ) {
    const x = Number.isFinite(Number(event.clientX)) ? event.clientX : job.x;
    const y = Number.isFinite(Number(event.clientY)) ? event.clientY : job.y;
    if (job.outsideCell || !pointIsInsideElementRect(job.cell, x, y)) {
      state.suppressNextCellClick = true;
    }
  }
  state.editAutoScroll = null;
  if (state.editAutoScrollFrame) {
    cancelAnimationFrame(state.editAutoScrollFrame);
    state.editAutoScrollFrame = null;
  }
}



function pointIsInsideElementRect(element, clientX, clientY) {
  if (!element?.isConnected) return false;
  const rect = element.getBoundingClientRect();
  return clientX >= rect.left && clientX <= rect.right && clientY >= rect.top && clientY <= rect.bottom;
}



function scheduleEditAutoScroll() {
  if (state.editAutoScrollFrame) return;
  state.editAutoScrollFrame = requestAnimationFrame(runEditAutoScroll);
}



function runEditAutoScroll() {
  state.editAutoScrollFrame = null;
  const job = state.editAutoScroll;
  if (!job?.cell?.isConnected) return;

  const cell = job.cell;
  const rect = cell.getBoundingClientRect();
  const threshold = Math.min(36, Math.max(16, rect.height / 3));
  let delta = 0;

  if (job.y > rect.bottom - threshold) {
    delta = Math.min(28, Math.max(1, job.y - (rect.bottom - threshold)) / 2);
  } else if (job.y < rect.top + threshold) {
    delta = -Math.min(28, Math.max(1, rect.top + threshold - job.y) / 2);
  }

  if (delta) {
    cell.scrollTop += delta;
  }
  scheduleEditAutoScroll();
}



function placeCaretAtEnd(element) {
  if (isTextInputElement(element)) {
    const end = String(element.value ?? "").length;
    element.focus?.({ preventScroll: true });
    element.setSelectionRange?.(end, end);
    return;
  }
  const selection = window.getSelection?.();
  if (!selection) return;
  const range = document.createRange();
  range.selectNodeContents(element);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
}



function placeCaretAtPoint(x, y, fallbackElement) {
  if (isTextInputElement(fallbackElement)) {
    fallbackElement.focus?.({ preventScroll: true });
    const end = String(fallbackElement.value ?? "").length;
    fallbackElement.setSelectionRange?.(end, end);
    return;
  }
  const selection = window.getSelection?.();
  if (!selection) return;

  let range = null;
  if (document.caretRangeFromPoint) {
    range = document.caretRangeFromPoint(x, y);
  } else if (document.caretPositionFromPoint) {
    const position = document.caretPositionFromPoint(x, y);
    if (position) {
      range = document.createRange();
      range.setStart(position.offsetNode, position.offset);
    }
  }

  if (!range || !fallbackElement.contains(range.startContainer)) {
    range = document.createRange();
    range.selectNodeContents(fallbackElement);
    range.collapse(false);
  }

  selection.removeAllRanges();
  selection.addRange(range);
}



function isTextInputElement(element) {
  return element instanceof HTMLTextAreaElement || element instanceof HTMLInputElement;
}



function insertLineBreakInEditable(element) {
  const selection = window.getSelection?.();
  if (!selection || !element) return false;
  let range = selection.rangeCount ? selection.getRangeAt(0) : null;
  if (
    !range ||
    !element.contains(range.startContainer) ||
    !element.contains(range.endContainer)
  ) {
    placeCaretAtEnd(element);
    range = selection.rangeCount ? selection.getRangeAt(0) : null;
  }
  if (!range) return false;

  range.deleteContents();
  const lineBreak = document.createTextNode("\n");
  const sentinel = document.createTextNode(CELL_EDIT_TRAILING_LINE_SENTINEL);
  const fragment = document.createDocumentFragment();
  fragment.appendChild(lineBreak);
  fragment.appendChild(sentinel);
  range.insertNode(fragment);
  range.setStartBefore(sentinel);
  range.collapse(true);
  selection.removeAllRanges();
  selection.addRange(range);
  try {
    element.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertLineBreak" }));
  } catch {
    element.dispatchEvent(new Event("input", { bubbles: true }));
  }
  return true;
}



function editablePlainText(element) {
  return stripCellEditSentinels(element?.textContent || "");
}



function stripCellEditSentinels(text) {
  return String(text ?? "").replaceAll(CELL_EDIT_TRAILING_LINE_SENTINEL, "");
}



function preserveCellEditSelectionForToolbar(event) {
  const editor = activeCellEditElementForFormatting();
  if (!editor || !cellEditToolbarTarget(event.target)) return;
  saveCellEditSelection(editor);
  $(editor).data("skipBlurCommit", true);
  if (!event.target.closest?.("select, input[type='color']")) {
    event.preventDefault();
  }
}



function cellEditToolbarTarget(target) {
  if (!target?.closest) return false;
  if (target.closest('[data-ribbon-command="bold"], [data-ribbon-command="italic"], [data-ribbon-underline-apply], [data-ribbon-underline-menu], [data-ribbon-underline-choice]')) return true;
  if (target.closest('[data-ribbon-style="fontFamily"], [data-ribbon-style="fontSize"], [data-ribbon-style="fontColor"]')) return true;
  if (target.closest('[data-ribbon-color-apply="fontColor"], [data-ribbon-color-menu="fontColor"], [data-ribbon-color-style="fontColor"], [data-ribbon-color-none="fontColor"], [data-ribbon-color-more="fontColor"]')) return true;
  return false;
}



function saveCellEditSelection(editor = activeCellEditElementForFormatting()) {
  if (!editor?.isConnected) return false;
  const selection = window.getSelection?.();
  if (!selection?.rangeCount) return false;
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.startContainer) || !editor.contains(range.endContainer)) return false;
  state.cellEditSelection = {
    editor,
    range: range.cloneRange(),
    start: cellEditTextOffsetForBoundary(editor, range.startContainer, range.startOffset),
    end: cellEditTextOffsetForBoundary(editor, range.endContainer, range.endOffset),
  };
  return true;
}



function restoreCellEditSelection(editor = activeCellEditElementForFormatting()) {
  if (!editor?.isConnected) return false;
  const saved = state.cellEditSelection;
  const selection = window.getSelection?.();
  if (!saved || saved.editor !== editor || !selection) return false;
  const range = document.createRange();
  const start = cellEditBoundaryForTextOffset(editor, saved.start);
  const end = cellEditBoundaryForTextOffset(editor, saved.end);
  try {
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
  } catch {
    return false;
  }
  selection.removeAllRanges();
  selection.addRange(range);
  return true;
}



function clearCellEditSelection(editor = null) {
  if (!editor || state.cellEditSelection?.editor === editor) {
    state.cellEditSelection = null;
  }
}



function cellEditSelectionOffsets(editor) {
  restoreCellEditSelection(editor);
  const selection = window.getSelection?.();
  const textLength = editablePlainText(editor).length;
  if (!selection?.rangeCount) return { start: textLength, end: textLength };
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.startContainer) || !editor.contains(range.endContainer)) {
    const saved = state.cellEditSelection;
    if (saved?.editor === editor) {
      return {
        start: Math.max(0, Math.min(textLength, Math.min(saved.start, saved.end))),
        end: Math.max(0, Math.min(textLength, Math.max(saved.start, saved.end))),
      };
    }
    return { start: textLength, end: textLength };
  }
  const start = cellEditTextOffsetForBoundary(editor, range.startContainer, range.startOffset);
  const end = cellEditTextOffsetForBoundary(editor, range.endContainer, range.endOffset);
  return {
    start: Math.max(0, Math.min(textLength, Math.min(start, end))),
    end: Math.max(0, Math.min(textLength, Math.max(start, end))),
  };
}



function cellEditTextOffsetForBoundary(editor, container, offset) {
  const range = document.createRange();
  range.selectNodeContents(editor);
  try {
    range.setEnd(container, offset);
    return stripCellEditSentinels(range.toString()).length;
  } catch {
    return editablePlainText(editor).length;
  }
}



function cellEditBoundaryForTextOffset(editor, targetOffset) {
  const target = Math.max(0, Math.min(editablePlainText(editor).length, Number(targetOffset) || 0));
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  let logicalOffset = 0;
  let last = { node: editor, offset: editor.childNodes.length };
  while (walker.nextNode()) {
    const node = walker.currentNode;
    const text = String(node.nodeValue || "");
    for (let index = 0; index <= text.length; index += 1) {
      if (logicalOffset === target) return { node, offset: index };
      if (index >= text.length) break;
      if (text.charAt(index) !== CELL_EDIT_TRAILING_LINE_SENTINEL) logicalOffset += 1;
    }
    last = { node, offset: text.length };
  }
  return last;
}

export {
  activateNameBoxTargets,
  activeSelectedImageIds,
  applyHeaderSelectionIndex,
  applyNameBoxText,
  applySelectionRange,
  applyValidationImeModeToEditable,
  beginCellEdit,
  beginCellEditAtPoint,
  beginNameBoxEdit,
  canvasAlignedTextLeft,
  canvasDevicePixelRatio,
  canvasFont,
  canvasFontDisplayScale,
  canvasFontSize,
  canvasLayerBounds,
  canvasLineBaseline,
  canvasTextLineHeight,
  cellAnnotationMarkersHtml,
  cellAnnotationTitle,
  cellAutoFilterButtonHtml,
  cellContentStyle,
  cellEditBoundaryForTextOffset,
  cellEditHtmlForCell,
  cellEditRunHtml,
  cellEditSelectionOffsets,
  cellEditTextOffsetForBoundary,
  cellEditToolbarTarget,
  cellHasComment,
  cellHasConditionalFormat,
  cellHasConstant,
  cellHasDomContent,
  cellHasFill,
  cellHasFontColor,
  cellHasFormula,
  cellHasKeyboardNavigationContent,
  cellHasNote,
  cellInnerHtml,
  cellNeedsDomOverlay,
  cellValidationDropdownButtonHtml,
  chartObjectOutlineStyle,
  clampElbowConnectorBendForPoints,
  clearCellEditSelection,
  clearPictureCropState,
  clipRangeToVisibleRange,
  clippedRangeFrameBorderStyle,
  closeNameBoxMenu,
  closeNameBoxMenuOnOutsidePointer,
  columnSelectionRange,
  coordinateToPixels,
  copiedRangeForActiveSheet,
  cssFontSizeToPoints,
  customPathEditablePointCount,
  customPathVertexForKey,
  dataRibbonForecastPointCount,
  dataRibbonRangeHasData,
  dataRibbonSelectionContext,
  defineNameFromNameBox,
  editableCellOverlayHeight,
  editableCellOverlayMetrics,
  editableCellPaddingTop,
  editableCellVerticalAlignment,
  editablePlainText,
  elbowConnectorAdjustmentHandlesForFrame,
  elbowConnectorUsesHorizontalMiddle,
  ensureCellElementForPoint,
  extendHeaderSelection,
  extendSelectionByKeyboard,
  finishNameBoxEdit,
  finishNameBoxEditOnBlur,
  firstVisibleSheetAxisIndex,
  fontOptionFromCss,
  fontSizeOptionFromCss,
  gridPositionHasFill,
  handleNameBoxKeydown,
  handleNameBoxMenuItemClick,
  handleSheetHostCellClick,
  headerSelectionAnchor,
  headerSelectionRange,
  imageBoxBaseStyle,
  imageBoxStyle,
  imageBoxToPixels,
  imageFlipH,
  imageFlipV,
  imageObjectHtml,
  imageRotationDegrees,
  imageSelectionBoundsToPixels,
  imageSelectionBoxToPixels,
  imageSelectionFrameStyle,
  imageViewportBounds,
  imageVisualBoundsToPixels,
  imageVisualBoxToPixels,
  inkStrokeSvg,
  insertLineBreakInEditable,
  isActiveImage,
  isBoldCssValue,
  isCopiedSourceImage,
  isCopiedSourceRange,
  isPictureCropModeActive,
  isSheetAxisIndexHidden,
  isTextInputElement,
  keyboardNavigationCellHasContent,
  keyboardNavigationEdgeTarget,
  keyboardNavigationFindEdgeInColumn,
  keyboardNavigationFindEdgeInRow,
  keyboardNavigationLastUsedColumnInRow,
  keyboardNavigationPageRowStep,
  keyboardNavigationTargetFromEvent,
  keyboardNavigationUsedRange,
  keyboardSelectionAdjacentTarget,
  keyboardSelectionExtensionTargetFromEvent,
  moveSelectionByKeyboard,
  nameBoxVisibleDefinitions,
  normalizeFontName,
  objectVisualStyle,
  openNameBoxMenu,
  outlineRangeHasLevel,
  overlayCellCss,
  paintCanvasPhonetics,
  paintCanvasStrikethrough,
  paintCanvasUnderline,
  parseCustomPathVertexKey,
  phoneticRunsForLine,
  pictureCropContentStyle,
  pictureOutlineCssStyle,
  pictureOutlineStyle,
  pixelBoxesIntersect,
  placeCaretAtEnd,
  placeCaretAtPoint,
  pointIsInsideElementRect,
  preserveCellEditSelectionForToolbar,
  rangesEqual,
  rawActiveSelectionRanges,
  rawRangesEqual,
  renderCellHtml,
  renderCopyFrame,
  renderFormulaReferenceFrame,
  renderGroupShapeHandles,
  renderImageGroupSelectionFrame,
  renderImageLayer,
  renderImageSelectionFrame,
  renderImageSelectionFrames,
  renderInkLayer,
  renderLineShapeHandles,
  renderPictureCropHandles,
  renderRangeFrameHtml,
  renderResizeHandles,
  renderSelectionFrame,
  renderShapeAdjustmentHandles,
  renderShapeResizeHandles,
  resolveNameBoxTargets,
  restoreCellEditSelection,
  rotatedPixelBox,
  roundCssNumber,
  rowSelectionRange,
  rtlColumnCoordinateToPixels,
  runEditAutoScroll,
  saveCellEditSelection,
  scheduleEditAutoScroll,
  scrollRangeIntoView,
  selectCell,
  selectCellByKeyboard,
  selectCellPoint,
  selectElementText,
  selectedCroppablePicture,
  selectedImageGroups,
  selectedImageObject,
  selectedImageObjectById,
  selectedImageObjects,
  selectedInkStrokeIdSet,
  selectedInkStrokes,
  selectionLabelText,
  selectionToolbarPoint,
  setImageElementRenderSource,
  setRibbonButtonActive,
  setRibbonCommandDisabled,
  setSelectedCell,
  setSelectedImageIds,
  shapeAdjustmentBoxFromFrame,
  shapeAdjustmentBoxFromShape,
  shapeAdjustmentHandles,
  shapeAdjustmentHandlesForFrame,
  shapeCustomPathHandles,
  shapeCustomPathIsClosed,
  shapeFilterValues,
  shapeTextLayerHtml,
  shapeTextLayerScaleStyle,
  shapeTextLayerStyle,
  shapeTextPlainLayerHtml,
  shapeTextPlainValue,
  shapeTextRunLayerStyle,
  shapeTextRunsLayerHtml,
  shapeVisualEffectStyle,
  sheetColumnsRightToLeft,
  shouldHideGridBottom,
  shouldHideGridRight,
  splitShapeTextLines,
  startCellEdit,
  startColumnHeaderSelection,
  startEditAutoScroll,
  startHeaderSelection,
  startRowHeaderSelection,
  startSelectedCellEdit,
  stopEditAutoScroll,
  stripCellEditSentinels,
  syncDataRibbonControls,
  syncEditableCellAlignment,
  syncPictureCropSelectionState,
  syncRibbonControlsForCell,
  unionImageBox,
  unionImageSelectionBounds,
  updateEditAutoScroll,
  updateFormulaBarForCell,
  updateFormulaBarForSelection,
  updateHeaderSelectionDrag,
  updateHeaderSelectionFromClient,
  visibleCellTarget,
  visibleSheetAxisIndex,
  visualGridColumnStart,
  worksheetGroupSelectionLayerStyle,
  worksheetImageRenderSource,
  worksheetObjectLayerIndex,
  worksheetObjectLayerStyle,
  worksheetObjectLayerZIndex,
  wrapCanvasText,
};
