// Feature module; shared bindings are imported explicitly.
import {
  $,
  $ribbonPanels,
  CELL_BORDER_SIDES,
  CWS_CLIPBOARD_DATA_MIME,
  CWS_CLIPBOARD_DATA_WEB_MIME,
  CWS_CLIPBOARD_HTML_MARKER,
  CWS_CLIPBOARD_RANGE_MARKER,
  CWS_CLIPBOARD_RANGE_MIME,
  DEFAULT_RECENT_SHAPE_TYPES,
  DEFAULT_SHAPE_TYPE,
  DEFAULT_THEME_COLORS,
  DRAW_INK_ERASERS,
  HTML_NAMED_COLORS,
  INSERT_SYMBOL_CHOICES,
  INTERNAL_CLIPBOARD_STORAGE_KEY,
  RIBBON_BORDER_PRESETS,
  RIBBON_COLOR_STYLE,
  RIBBON_COLOR_TINTS,
  RIBBON_SHAPE_GROUPS,
  RIBBON_STANDARD_COLORS,
  RIBBON_UNDERLINE_PRESETS,
  SHAPE_ARROW_PRESETS,
  SHAPE_ARROW_SIZE_PRESETS,
  SHAPE_ARROW_TYPE_PRESETS,
  SHAPE_FILL_PATTERN_PRESETS,
  SHAPE_FIT_INSET,
  SHAPE_GLOW_PRESETS,
  SHAPE_LINE_CAP_PRESETS,
  SHAPE_LINE_COMPOUND_PRESETS,
  SHAPE_LINE_DASH_PRESETS,
  SHAPE_LINE_JOIN_PRESETS,
  SHAPE_OUTLINE_WEIGHTS,
  SHAPE_REFLECTION_PRESETS,
  SHAPE_SOFT_EDGE_PRESETS,
  SHAPE_STYLE_PRESETS,
  createIcons,
  icons,
  state,
} from "../../app-core.js";
import {
  escapeRegExp,
} from "../data/filter-and-text-tools.js";
import {
  beginShapeDrawingMode,
} from "../drawing/charts-and-insertion.js";
import {
  clamp01,
  clampNumber,
  insertionPointForObjects,
  isLineShapeObject,
  normalizeShapeArrowSize,
  normalizeShapeArrowType,
  normalizeShapeFillPattern,
  normalizeShapeGlow,
  normalizeShapeLineCap,
  normalizeShapeLineCompound,
  normalizeShapeLineDash,
  normalizeShapeLineJoin,
  normalizeShapeOutlineWidth,
  normalizeShapeReflection,
  normalizeShapeSoftEdges,
  normalizeShapeTextMargin,
  normalizeShapeTextOutlineWidth,
  shapeDescriptorForType,
  shapePresetForType,
  shapeSvgMarkup,
} from "../drawing/shape-rendering.js";
import {
  alignSelectedShapes,
  applySelectedShapeRotation,
  applySelectedShapeSize,
  applySelectedShapeStyle,
  applyShapeArrowPreset,
  applyShapeArrowSizePreset,
  applyShapeArrowTypePreset,
  applyShapeRotateAction,
  reorderSelectedObjects,
} from "../drawing/transforms-and-connectors.js";
import {
  applyCellEditCssPatchToSelection,
  applyCellEditUnderlinePreset,
  deleteSelectedImages,
  groupedShapeObjectTargets,
  normalizeCellEditColor,
  openShapeTextEditor,
} from "./autofill.js";
import {
  isElementImeCompositionEvent,
  selectedObjectSummary,
} from "./keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSheetGroupEditIndexes,
  coerceUserInput,
  commitCellValue,
  cssEscape,
  getCellRawInput,
  restoreGroupedSheetsAfterFailure,
  selectedObjectFormatMode,
  syncShapeFormatControls,
  updateSelectionUi,
} from "./pointers-and-formula-evaluation.js";
import {
  activeSelectedImageIds,
  fontOptionFromCss,
  fontSizeOptionFromCss,
  imageBoxToPixels,
  imageRotationDegrees,
  roundCssNumber,
  selectedImageGroups,
  selectedImageObject,
  selectedImageObjectById,
  selectedImageObjects,
  selectionToolbarPoint,
  setSelectedImageIds,
  unionImageBox,
} from "./selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneImages,
  cloneJsonLikeValue,
  pushHistory,
} from "./structure-and-history.js";
import {
  applyBorderPreset,
  applyCssToSelection,
  applyNumberFormatToSelection,
  applyShapeTextCssPatchToSelection,
  numberFormatForRibbon,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  clipboardCellEntries,
  closeCellContextMenu,
  closeCellStructureDialog,
  closeFormatCellsDialog,
  closeLinkDialog,
  closeNameDialog,
  formatFontOptions,
  formatOption,
  openLinkDialog,
  openObjectLinkDialog,
} from "../ui/annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  escapeAttr,
  escapeHtml,
  setStatus,
} from "../ui/help-and-utilities.js";
import {
  beginInkEraserMode,
  commandMenuEscapeBlockedByDialog,
  consumeCommandMenuKey,
  enhanceCommandMenuAccessKeys,
} from "../ui/printing-and-home-tools.js";
import {
  activeCellEditElementForFormatting,
  closeFileSideMenu,
  renameObjectForSheetGroup,
  renameSelectedObjects,
  switchRibbonTab,
} from "../ui/ribbon-and-links.js";
import {
  closeSheetTabContextMenu,
  ensureSheetProtectionPermission,
  objectFormatPaneTitle,
} from "../ui/sheet-protection-and-translation.js";
import {
  applyTint,
  fontFamilyForExcel,
  verticalAlignToFlex,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  activeSheet,
} from "../workbook/saving-and-export.js";


function squaredDistance(a, b) {
  const dx = Number(a.x) - Number(b.x);
  const dy = Number(a.y) - Number(b.y);
  return dx * dx + dy * dy;
}



function outlinePointsToPathD(points) {
  const [first, ...rest] = points;
  if (!first) return "";
  return `M${first.x} ${first.y} ${rest.map((point) => `L${point.x} ${point.y}`).join(" ")} Z`;
}



function roundPoint(point) {
  return { x: roundSvgNumber(point.x), y: roundSvgNumber(point.y) };
}



function removeAdjacentDuplicatePoints(points) {
  const normalized = [];
  points.forEach((point) => {
    const previous = normalized[normalized.length - 1];
    if (!previous || Math.abs(previous.x - point.x) > 0.001 || Math.abs(previous.y - point.y) > 0.001) normalized.push(point);
  });
  const first = normalized[0];
  const last = normalized[normalized.length - 1];
  if (first && last && Math.abs(first.x - last.x) < 0.001 && Math.abs(first.y - last.y) < 0.001) normalized.pop();
  return normalized;
}



function cloudOutlinePoints(body, stepsPerCurve = 18) {
  const w = body.right - body.left;
  const h = body.bottom - body.top;
  const p = (x, y) => ({ x: body.left + w * x, y: body.top + h * y });
  const curves = [
    [p(0.02, 0.65), p(0.02, 0.42), p(0.11, 0.32), p(0.2, 0.32)],
    [p(0.2, 0.32), p(0.22, 0.1), p(0.4, 0.05), p(0.48, 0.22)],
    [p(0.48, 0.22), p(0.58, 0.02), p(0.78, 0.14), p(0.8, 0.35)],
    [p(0.8, 0.35), p(0.93, 0.35), p(0.99, 0.48), p(0.99, 0.65)],
    [p(0.99, 0.65), p(0.99, 0.84), p(0.85, 0.98), p(0.68, 0.9)],
    [p(0.68, 0.9), p(0.55, 1.02), p(0.36, 0.98), p(0.31, 0.84)],
    [p(0.31, 0.84), p(0.17, 0.91), p(0.02, 0.82), p(0.02, 0.65)],
  ];
  const points = [];
  curves.forEach((curve, curveIndex) => {
    for (let step = curveIndex === 0 ? 0 : 1; step <= stepsPerCurve; step += 1) {
      const point = cubicBezierPoint(curve[0], curve[1], curve[2], curve[3], step / stepsPerCurve);
      points.push({ x: roundSvgNumber(point.x), y: roundSvgNumber(point.y) });
    }
  });
  return points;
}



function cubicBezierPoint(p0, p1, p2, p3, t) {
  const mt = 1 - t;
  const mt2 = mt * mt;
  const t2 = t * t;
  return {
    x: mt2 * mt * p0.x + 3 * mt2 * t * p1.x + 3 * mt * t2 * p2.x + t2 * t * p3.x,
    y: mt2 * mt * p0.y + 3 * mt2 * t * p1.y + 3 * mt * t2 * p2.y + t2 * t * p3.y,
  };
}



function segmentIntersectionPoint(a, b, c, d) {
  const r = { x: b.x - a.x, y: b.y - a.y };
  const s = { x: d.x - c.x, y: d.y - c.y };
  const denominator = r.x * s.y - r.y * s.x;
  if (Math.abs(denominator) < 0.000001) return null;
  const q = { x: c.x - a.x, y: c.y - a.y };
  const t = (q.x * s.y - q.y * s.x) / denominator;
  const u = (q.x * r.y - q.y * r.x) / denominator;
  if (t < -0.000001 || t > 1.000001 || u < -0.000001 || u > 1.000001) return null;
  return {
    t,
    u,
    x: a.x + r.x * t,
    y: a.y + r.y * t,
  };
}



function calloutRectPathD(geometry) {
  const [first, ...rest] = geometry.points || [];
  if (!first) return "";
  return `M${first.x} ${first.y} ${rest.map((point) => `L${point.x} ${point.y}`).join(" ")} Z`;
}



function cloudPathD(body) {
  const w = body.right - body.left;
  const h = body.bottom - body.top;
  const x = (pct) => roundSvgNumber(body.left + w * pct);
  const y = (pct) => roundSvgNumber(body.top + h * pct);
  return `M${x(0.02)} ${y(0.65)} C${x(0.02)} ${y(0.42)} ${x(0.11)} ${y(0.32)} ${x(0.2)} ${y(0.32)} C${x(0.22)} ${y(0.1)} ${x(0.4)} ${y(0.05)} ${x(0.48)} ${y(0.22)} C${x(0.58)} ${y(0.02)} ${x(0.78)} ${y(0.14)} ${x(0.8)} ${y(0.35)} C${x(0.93)} ${y(0.35)} ${x(0.99)} ${y(0.48)} ${x(0.99)} ${y(0.65)} C${x(0.99)} ${y(0.84)} ${x(0.85)} ${y(0.98)} ${x(0.68)} ${y(0.9)} C${x(0.55)} ${y(1.02)} ${x(0.36)} ${y(0.98)} ${x(0.31)} ${y(0.84)} C${x(0.17)} ${y(0.91)} ${x(0.02)} ${y(0.82)} ${x(0.02)} ${y(0.65)} Z`;
}



function fitSvgPointsString(points, inset = SHAPE_FIT_INSET) {
  const normalized = points.map(([x, y]) => [Number(x), Number(y)]).filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
  if (!normalized.length) return "";
  const xs = normalized.map(([x]) => x);
  const ys = normalized.map(([, y]) => y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  const targetWidth = 100 - inset * 2;
  const targetHeight = 60 - inset * 2;
  const sourceWidth = Math.max(0.0001, maxX - minX);
  const sourceHeight = Math.max(0.0001, maxY - minY);
  return svgPointsString(
    normalized.map(([x, y]) => [
      inset + ((x - minX) / sourceWidth) * targetWidth,
      inset + ((y - minY) / sourceHeight) * targetHeight,
    ]),
  );
}



function svgPointsString(points) {
  return points.map(([x, y]) => `${roundSvgNumber(x)},${roundSvgNumber(y)}`).join(" ");
}



function roundSvgNumber(value) {
  return Math.round(value * 10) / 10;
}



function svgPointFromLinePoint(x, y) {
  return {
    x: roundSvgNumber(6 + clamp01(Number(x)) * 88),
    y: roundSvgNumber(6 + clamp01(Number(y)) * 48),
  };
}



function defaultTextBoxDescriptor() {
  return shapeDescriptorForType("textBox");
}



function svgDataUrl(svg) {
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}



function openLinkDialogForSelection() {
  const point = selectionToolbarPoint() || insertionPointForObjects();
  if (!point) {
    setStatus("リンクを設定するセルを選択してください。");
    return;
  }
  openLinkDialog(point.row, point.col);
}



function openSymbolDialog() {
  const point = selectionToolbarPoint() || insertionPointForObjects();
  if (!point) {
    setStatus("記号を挿入するセルを選択してください。");
    return;
  }
  closeFormatCellsDialog();
  closeCellStructureDialog();
  closeNameDialog();
  closeLinkDialog();
  closeSymbolDialog({ restoreFocus: false });

  state.symbolDialog = {
    previousFocus: document.activeElement,
    point,
  };
  const overlay = document.createElement("div");
  overlay.id = "symbolDialogOverlay";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = symbolDialogHtml();
  document.body.append(overlay);
  overlay.querySelector("[data-symbol-choice]")?.focus({ preventScroll: true });
}



function symbolDialogHtml() {
  const choices = INSERT_SYMBOL_CHOICES.map(
    (symbol) => `<button class="symbol-choice" type="button" data-symbol-choice="${escapeAttr(symbol)}">${escapeHtml(symbol)}</button>`,
  ).join("");
  return `
    <div id="symbolDialog" class="excel-mini-dialog symbol-dialog" role="dialog" aria-modal="true" aria-labelledby="symbolDialogTitle">
      <div class="excel-mini-titlebar" id="symbolDialogTitle">記号と特殊文字</div>
      <div class="excel-mini-body">
        <div class="symbol-dialog-grid" role="grid" aria-label="記号">${choices}</div>
      </div>
      <div class="excel-mini-actions">
        <span class="excel-mini-action-spacer"></span>
        <button class="excel-mini-button" type="button" data-symbol-dialog-action="cancel">キャンセル</button>
      </div>
    </div>
  `;
}



function handleSymbolChoice(event) {
  const symbol = event.currentTarget.dataset.symbolChoice || "";
  if (!symbol || !state.symbolDialog) return;
  const point = state.symbolDialog.point || selectionToolbarPoint();
  closeSymbolDialog({ restoreFocus: false });
  if (!point) return;
  const sheet = activeSheet();
  const current = getCellRawInput(sheet, point.row, point.col);
  commitCellValue(point.row, point.col, `${current}${symbol}`);
}



function handleSymbolDialogAction(event) {
  if (event.currentTarget.dataset.symbolDialogAction === "cancel") closeSymbolDialog();
}



function closeSymbolDialog(options = {}) {
  document.querySelector("#symbolDialogOverlay")?.remove();
  const previousFocus = state.symbolDialog?.previousFocus;
  state.symbolDialog = null;
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) {
    previousFocus.focus?.({ preventScroll: true });
  }
}



function closeSymbolDialogOnEscape(event) {
  if (event.key !== "Escape" || !state.symbolDialog) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  closeSymbolDialog();
}



function uniqueWorksheetObjectId(prefix) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}



function handleRibbonStyleChange(control) {
  if (!state.model) {
    setStatus("先に Excel ファイルを開いてください。");
    return;
  }

  const style = control.dataset.ribbonStyle;
  const value = normalizeColorValue(control.value) || control.value;
  switch (style) {
    case "fontFamily":
      if (applyCellEditCssPatchToSelection({ fontFamily: fontFamilyForExcel(value) }, `フォント：${value}`)) break;
      applyCssToSelection({ fontFamily: fontFamilyForExcel(value) }, `フォント：${value}`);
      break;
    case "fontSize":
      if (applyCellEditCssPatchToSelection({ fontSize: `${value}pt` }, `フォント サイズ：${value}`)) break;
      applyCssToSelection({ fontSize: `${value}pt` }, `フォント サイズ：${value}`);
      break;
    case "fillColor":
      setRibbonColorValue(style, value);
      applyCssToSelection({ backgroundColor: value }, `塗りつぶしの色：${value}`);
      break;
    case "fontColor":
      setRibbonColorValue(style, value);
      if (applyCellEditCssPatchToSelection({ color: value }, `フォントの色：${value}`)) break;
      applyCssToSelection({ color: value }, `フォントの色：${value}`);
      break;
    case "numberFormat":
      applyNumberFormatToSelection(numberFormatForRibbon(value), `表示形式：${value}`);
      break;
    default:
      break;
  }
}



function applyRibbonColor(style) {
  const control = ribbonColorInput(style);
  if (!control) return;
  handleRibbonStyleChange(control);
}



function applyRibbonBorder(event) {
  event?.preventDefault?.();
  const preset = document.querySelector("[data-ribbon-border-apply]")?.dataset.ribbonBorderCurrent || "outside";
  applyBorderPreset(preset);
}



function applyRibbonUnderline(event) {
  event?.preventDefault?.();
  const preset = document.querySelector("[data-ribbon-underline-apply]")?.dataset.ribbonUnderlineCurrent || "underline";
  applyUnderlinePreset(preset, { toggle: true });
}



function applyUnderlinePreset(preset, options = {}) {
  const selectedPreset = underlinePresetForValue(preset) || RIBBON_UNDERLINE_PRESETS[0];
  if (applyCellEditUnderlinePreset(selectedPreset.value, options)) return;
  const range = activeSelectionRange();
  if (!range) {
    const image = selectedImageObjects().find((item) => item?.shape && !isLineShapeObject(item));
    if (image) {
      const currentPreset = underlinePresetForValue(image.shape?.textUnderline);
      const nextValue = options.toggle && currentPreset?.value === selectedPreset.value ? null : selectedPreset.value;
      setRibbonUnderlinePreset(selectedPreset.value);
      applyShapeTextCssPatchToSelection({ textDecoration: nextValue }, nextValue ? selectedPreset.label.replace(/\(.+\)/, "") : "下線なし");
      return;
    }
    setStatus("セルを選択してください。");
    return;
  }
  const sheet = activeSheet();
  const point = selectionToolbarPoint() || { row: range.top, col: range.left };
  const activeCell = sheet.cells[cellKey(point.row, point.col)];
  const currentPreset = underlinePresetForTextDecoration(activeCell?.css?.textDecoration);
  const nextValue = options.toggle && currentPreset?.value === selectedPreset.value ? null : selectedPreset.value;
  setRibbonUnderlinePreset(selectedPreset.value);
  applyCssToSelection({ textDecoration: nextValue }, nextValue ? selectedPreset.label.replace(/\(.+\)/, "") : "下線なし");
}



function openRibbonUnderlineMenu(event) {
  event?.preventDefault?.();
  const button = event.currentTarget;

  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeRibbonBorderMenu();
  closeRibbonShapeMenu();
  closeRibbonInkEraserMenu();
  closeShapeFormatMenu();

  const current = document.querySelector("[data-ribbon-underline-apply]")?.dataset.ribbonUnderlineCurrent || "underline";
  $("#ribbonUnderlineMenu").remove();
  $("body").append(ribbonUnderlineMenuHtml(current));

  const menu = document.querySelector("#ribbonUnderlineMenu");
  if (!menu) return;
  enhanceCommandMenuAccessKeys(menu);
  const rect = button.getBoundingClientRect();
  const top = Math.min(window.innerHeight - menu.offsetHeight - 8, rect.bottom + 2);
  const left = Math.min(window.innerWidth - menu.offsetWidth - 8, Math.max(8, rect.left - menu.offsetWidth + rect.width));
  menu.style.top = `${Math.max(8, Math.round(top))}px`;
  menu.style.left = `${Math.max(8, Math.round(left))}px`;
  state.ribbonUnderlineMenu = true;
  if (!activeCellEditElementForFormatting()) {
    menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
  }
}



function closeRibbonUnderlineMenu() {
  $("#ribbonUnderlineMenu").remove();
  state.ribbonUnderlineMenu = null;
}



function closeRibbonUnderlineMenuOnOutsidePointer(event) {
  if (!state.ribbonUnderlineMenu) return;
  if (event.target.closest?.("#ribbonUnderlineMenu") || event.target.closest?.(".ribbon-underline-split")) return;
  closeRibbonUnderlineMenu();
}



function closeRibbonUnderlineMenuOnEscape(event) {
  if (event.key === "Escape" && state.ribbonUnderlineMenu) {
    if (commandMenuEscapeBlockedByDialog()) return;
    consumeCommandMenuKey(event);
    closeRibbonUnderlineMenu();
  }
}



function handleRibbonUnderlineChoice(event) {
  const preset = event.currentTarget.dataset.ribbonUnderlineChoice || "underline";
  setRibbonUnderlinePreset(preset);
  closeRibbonUnderlineMenu();
  applyUnderlinePreset(preset);
}



function openRibbonShapeMenu(event) {
  event?.preventDefault?.();
  event?.stopPropagation?.();
  const button = event.currentTarget;

  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonInkEraserMenu();
  closeShapeFormatMenu();

  $("#ribbonShapeMenu").remove();
  $("body").append(ribbonShapeMenuHtml());

  const menu = document.querySelector("#ribbonShapeMenu");
  if (!menu) return;
  enhanceCommandMenuAccessKeys(menu);
  const rect = button.getBoundingClientRect();
  const top = Math.min(window.innerHeight - menu.offsetHeight - 8, rect.bottom + 2);
  const left = Math.min(window.innerWidth - menu.offsetWidth - 8, Math.max(8, rect.left - 8));
  menu.style.top = `${Math.max(8, Math.round(top))}px`;
  menu.style.left = `${Math.max(8, Math.round(left))}px`;
  state.ribbonShapeMenu = true;
  menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
}



function closeRibbonShapeMenu() {
  $("#ribbonShapeMenu").remove();
  state.ribbonShapeMenu = null;
}



function closeRibbonShapeMenuOnOutsidePointer(event) {
  if (!state.ribbonShapeMenu) return;
  if (event.target.closest?.("#ribbonShapeMenu") || event.target.closest?.(".ribbon-shape-menu-button")) return;
  closeRibbonShapeMenu();
}



function closeRibbonShapeMenuOnEscape(event) {
  if (event.key === "Escape" && state.ribbonShapeMenu) {
    if (commandMenuEscapeBlockedByDialog()) return;
    consumeCommandMenuKey(event);
    closeRibbonShapeMenu();
  }
}



function handleRibbonShapeChoice(event) {
  const type = event.currentTarget.dataset.ribbonShapeChoice || DEFAULT_SHAPE_TYPE;
  closeRibbonShapeMenu();
  beginShapeDrawingMode(type);
}



function openRibbonInkEraserMenu(event) {
  event?.preventDefault?.();
  event?.stopPropagation?.();
  const button = event?.currentTarget || event;
  if (!button) return;

  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeRibbonInkEraserMenu();
  closeShapeFormatMenu();

  $("#ribbonInkEraserMenu").remove();
  $("body").append(ribbonInkEraserMenuHtml());

  const menu = document.querySelector("#ribbonInkEraserMenu");
  if (!menu) return;
  enhanceCommandMenuAccessKeys(menu);
  const rect = button.getBoundingClientRect();
  const top = Math.min(window.innerHeight - menu.offsetHeight - 8, rect.bottom + 2);
  const left = Math.min(window.innerWidth - menu.offsetWidth - 8, Math.max(8, rect.left));
  menu.style.top = `${Math.max(8, Math.round(top))}px`;
  menu.style.left = `${Math.max(8, Math.round(left))}px`;
  state.ribbonInkEraserMenu = true;
  menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
}



function closeRibbonInkEraserMenu() {
  $("#ribbonInkEraserMenu").remove();
  state.ribbonInkEraserMenu = null;
}



function closeRibbonInkEraserMenuOnOutsidePointer(event) {
  if (!state.ribbonInkEraserMenu) return;
  if (event.target.closest?.("#ribbonInkEraserMenu") || event.target.closest?.(".ribbon-ink-eraser-button")) return;
  closeRibbonInkEraserMenu();
}



function closeRibbonInkEraserMenuOnEscape(event) {
  if (event.key === "Escape" && state.ribbonInkEraserMenu) {
    if (commandMenuEscapeBlockedByDialog()) return;
    consumeCommandMenuKey(event);
    closeRibbonInkEraserMenu();
  }
}



function handleRibbonInkEraserChoice(event) {
  const eraserKey = event.currentTarget.dataset.ribbonInkEraserChoice || "stroke";
  closeRibbonInkEraserMenu();
  beginInkEraserMode(eraserKey);
}



function ribbonInkEraserMenuHtml() {
  const current = state.currentInkEraser || "stroke";
  const buttons = Object.entries(DRAW_INK_ERASERS).map(([key, preset]) => {
    const active = key === current ? " is-active" : "";
    return `<button class="ribbon-ink-eraser-choice${active}" type="button" role="menuitem" data-ribbon-ink-eraser-choice="${escapeAttr(key)}">
      <span class="ribbon-ink-eraser-mark" data-ink-eraser-mark="${escapeAttr(key)}" aria-hidden="true"></span>
      <span>${accessKeyLabelHtml(preset.menuLabel || preset.label)}</span>
    </button>`;
  }).join("");
  return `<div id="ribbonInkEraserMenu" class="ribbon-ink-eraser-menu" role="menu" aria-label="消しゴム">${buttons}</div>`;
}



function openShapeFormatMenu(event, type) {
  event?.preventDefault?.();
  event?.stopPropagation?.();
  const button = event.currentTarget;
  if (!button) return;
  if (button.disabled) return;

  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeRibbonInkEraserMenu();
  closeShapeFormatMenu();

  $("body").append(shapeFormatMenuHtml(type));
  const menu = document.querySelector("#shapeFormatMenu");
  if (!menu) return;
  enhanceCommandMenuAccessKeys(menu);
  const rect = button.getBoundingClientRect();
  const top = Math.min(window.innerHeight - menu.offsetHeight - 8, rect.bottom + 2);
  const left = Math.min(window.innerWidth - menu.offsetWidth - 8, Math.max(8, rect.left));
  menu.style.top = `${Math.max(8, Math.round(top))}px`;
  menu.style.left = `${Math.max(8, Math.round(left))}px`;
  state.ribbonShapeFormatMenu = { type };
  menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
}



function closeShapeFormatMenu() {
  $("#shapeFormatMenu").remove();
  state.ribbonShapeFormatMenu = null;
}



function closeShapeFormatMenuOnOutsidePointer(event) {
  if (!state.ribbonShapeFormatMenu) return;
  if (event.target.closest?.("#shapeFormatMenu") || event.target.closest?.("[data-shape-fill-menu], [data-shape-outline-menu], [data-shape-effects-menu], [data-shape-text-fill-menu], [data-shape-text-outline-menu], [data-shape-align-menu], [data-shape-rotate-menu]")) return;
  closeShapeFormatMenu();
}



function closeShapeFormatMenuOnEscape(event) {
  if (event.key === "Escape" && state.ribbonShapeFormatMenu) {
    if (commandMenuEscapeBlockedByDialog()) return;
    consumeCommandMenuKey(event);
    closeShapeFormatMenu();
  }
}



function handleShapeStylePresetClick(event) {
  const preset = SHAPE_STYLE_PRESETS[event.currentTarget.dataset.shapeStylePreset];
  if (!preset) return;
  applySelectedShapeStyle(preset, "図形のスタイル");
}



function handleShapeFormatMenuAction(event) {
  const button = event.currentTarget;
  const action = button.dataset.shapeFormatAction || "";
  const value = button.dataset.shapeFormatValue || "";
  closeShapeFormatMenu();
  switch (action) {
    case "fill-color":
      applySelectedShapeStyle({ fillColor: normalizeColorValue(value) }, "図形の塗りつぶし");
      break;
    case "fill-none":
      applySelectedShapeStyle({ fillColor: "none" }, "塗りつぶしなし");
      break;
    case "fill-pattern":
      applySelectedShapeStyle({
        fillPattern: normalizeShapeFillPattern({
          preset: value,
          foreground: selectedImageObject()?.shape?.outlineColor || "#4472c4",
          background: selectedImageObject()?.shape?.fillColor || "rgba(255,255,255,0)",
        }),
      }, "パターン");
      break;
    case "outline-color":
      applySelectedShapeStyle({ outlineColor: normalizeColorValue(value) }, "図形の枠線");
      break;
    case "outline-none":
      applySelectedShapeStyle({ outlineColor: "none" }, "枠線なし");
      break;
    case "outline-weight":
      applySelectedShapeStyle({ outlineWidth: normalizeShapeOutlineWidth(value) }, `${value} pt`);
      break;
    case "outline-dash":
      applySelectedShapeStyle({ lineDash: normalizeShapeLineDash(value) }, "線の種類");
      break;
    case "outline-compound":
      applySelectedShapeStyle({ lineCompound: normalizeShapeLineCompound(value) }, "線の一重/多重");
      break;
    case "outline-cap":
      applySelectedShapeStyle({ lineCap: normalizeShapeLineCap(value) }, "線の先端");
      break;
    case "outline-join":
      applySelectedShapeStyle({ lineJoin: normalizeShapeLineJoin(value) }, "線の結合点");
      break;
    case "text-fill-color":
      applySelectedShapeStyle({ textColor: normalizeColorValue(value) || "#1f2937" }, "文字の塗りつぶし");
      break;
    case "text-outline-color":
      applySelectedShapeStyle({ textOutlineColor: normalizeColorValue(value) || "#1f2937", textOutlineWidth: 0.75 }, "文字の輪郭");
      break;
    case "text-outline-none":
      applySelectedShapeStyle({ textOutlineColor: "none" }, "文字の輪郭なし");
      break;
    case "arrow":
      applyShapeArrowPreset(value);
      break;
    case "arrow-start-type":
      applyShapeArrowTypePreset("start", value);
      break;
    case "arrow-end-type":
      applyShapeArrowTypePreset("end", value);
      break;
    case "arrow-start-size":
      applyShapeArrowSizePreset("start", value);
      break;
    case "arrow-end-size":
      applyShapeArrowSizePreset("end", value);
      break;
    case "shadow":
      applySelectedShapeStyle({ shadow: value === "none" ? "" : value }, value === "none" ? "影なし" : "影");
      break;
    case "glow":
      applySelectedShapeStyle({ glow: normalizeShapeGlow(value) }, value ? "光彩" : "光彩なし");
      break;
    case "soft-edge":
      applySelectedShapeStyle({ softEdges: normalizeShapeSoftEdges(value) }, Number(value) ? "ぼかし" : "ぼかしなし");
      break;
    case "reflection":
      applySelectedShapeStyle({ reflection: normalizeShapeReflection(value) }, value ? "反射" : "反射なし");
      break;
    case "align":
      alignSelectedShapes(value);
      break;
    case "rotate":
      applyShapeRotateAction(value);
      break;
    default:
      break;
  }
}



function shapeFormatMenuHtml(type) {
  const mode = selectedObjectFormatMode();
  const pictureMode = mode === "picture";
  const chartMode = mode === "chart";
  const title = {
    fill: chartMode ? "グラフ エリアの塗りつぶし" : pictureMode ? "図の塗りつぶし" : "図形の塗りつぶし",
    outline: chartMode ? "グラフの枠線" : pictureMode ? "図の枠線" : "図形の枠線",
    effects: chartMode ? "グラフの効果" : pictureMode ? "図の効果" : "図形の効果",
    textFill: "文字の塗りつぶし",
    textOutline: "文字の輪郭",
    align: "配置",
    rotate: "回転",
  }[type] || "図形";
  return `<div id="shapeFormatMenu" class="ribbon-shape-format-menu" role="menu" aria-label="${escapeAttr(title)}">${shapeFormatMenuBodyHtml(type)}</div>`;
}



function shapeFormatMenuBodyHtml(type) {
  if (type === "fill") return shapeFillMenuHtml();
  if (type === "outline") return shapeOutlineMenuHtml();
  if (type === "effects") return shapeEffectsMenuHtml();
  if (type === "textFill") return shapeTextFillMenuHtml();
  if (type === "textOutline") return shapeTextOutlineMenuHtml();
  if (type === "align") return shapeAlignMenuHtml();
  if (type === "rotate") return shapeRotateMenuHtml();
  return "";
}



function shapeFillMenuHtml() {
  const patterns = SHAPE_FILL_PATTERN_PRESETS
    .filter((preset) => preset.value)
    .map((preset) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="fill-pattern" data-shape-format-value="${escapeAttr(preset.value)}"><span class="shape-pattern-preview is-${escapeAttr(preset.value)}"></span><span>${escapeHtml(preset.label)}</span></button>`)
    .join("");
  return `${shapeColorSectionHtml("fill-color")}
    <button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="fill-none">${accessKeyLabelHtml("塗りつぶしなし(N)")}</button>
    <div class="shape-format-menu-separator"></div>
    <div class="shape-format-menu-title">パターン</div>${patterns}`;
}



function shapeOutlineMenuHtml() {
  const weights = SHAPE_OUTLINE_WEIGHTS.map((weight) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="outline-weight" data-shape-format-value="${weight}"><span class="shape-line-preview" style="--shape-line-width:${weight}px"></span><span>${weight} pt</span></button>`).join("");
  const dashes = SHAPE_LINE_DASH_PRESETS.map((preset) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="outline-dash" data-shape-format-value="${escapeAttr(preset.value)}"><span class="shape-line-preview is-${escapeAttr(preset.value)}"></span><span>${escapeHtml(preset.label)}</span></button>`).join("");
  const compounds = SHAPE_LINE_COMPOUND_PRESETS.map((preset) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="outline-compound" data-shape-format-value="${escapeAttr(preset.value)}"><span class="shape-line-preview is-compound-${escapeAttr(preset.value)}"></span><span>${escapeHtml(preset.label)}</span></button>`).join("");
  const caps = SHAPE_LINE_CAP_PRESETS.map((preset) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="outline-cap" data-shape-format-value="${escapeAttr(preset.value)}"><span class="shape-line-preview is-cap-${escapeAttr(preset.value)}"></span><span>${escapeHtml(preset.label)}</span></button>`).join("");
  const joins = SHAPE_LINE_JOIN_PRESETS.map((preset) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="outline-join" data-shape-format-value="${escapeAttr(preset.value)}"><span class="shape-line-preview is-join-${escapeAttr(preset.value)}"></span><span>${escapeHtml(preset.label)}</span></button>`).join("");
  const arrows = SHAPE_ARROW_PRESETS.map((preset) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="arrow" data-shape-format-value="${escapeAttr(preset.value)}"><span class="shape-arrow-preview is-${escapeAttr(preset.value)}"></span><span>${escapeHtml(preset.label)}</span></button>`).join("");
  return `${shapeColorSectionHtml("outline-color")}
    <button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="outline-none">${accessKeyLabelHtml("線なし(N)")}</button>
    <div class="shape-format-menu-separator"></div>
    <div class="shape-format-menu-title">太さ</div>${weights}
    <div class="shape-format-menu-title">一重線/多重線</div>${compounds}
    <div class="shape-format-menu-title">実線/点線</div>${dashes}
    <div class="shape-format-menu-title">線の先端</div>${caps}
    <div class="shape-format-menu-title">線の結合点</div>${joins}
    <div class="shape-format-menu-title">矢印</div>${arrows}
    <div class="shape-format-menu-title">始点矢印の種類</div>${shapeArrowTypeMenuItemsHtml("start")}
    <div class="shape-format-menu-title">終点矢印の種類</div>${shapeArrowTypeMenuItemsHtml("end")}
    <div class="shape-format-menu-title">始点矢印のサイズ</div>${shapeArrowSizeMenuItemsHtml("start")}
    <div class="shape-format-menu-title">終点矢印のサイズ</div>${shapeArrowSizeMenuItemsHtml("end")}`;
}



function shapeArrowTypeMenuItemsHtml(side) {
  const action = side === "start" ? "arrow-start-type" : "arrow-end-type";
  const sideClass = side === "start" ? "start" : "end";
  return SHAPE_ARROW_TYPE_PRESETS.map((preset) => {
    const previewType = preset.value === "none" ? "none" : normalizeShapeArrowType(preset.value) || "triangle";
    return `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="${action}" data-shape-format-value="${escapeAttr(preset.value)}"><span class="shape-arrow-preview is-${sideClass} is-arrow-${escapeAttr(previewType)}"></span><span>${escapeHtml(preset.label)}</span></button>`;
  }).join("");
}



function shapeArrowSizeMenuItemsHtml(side) {
  const action = side === "start" ? "arrow-start-size" : "arrow-end-size";
  const sideClass = side === "start" ? "start" : "end";
  return SHAPE_ARROW_SIZE_PRESETS.map((preset) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="${action}" data-shape-format-value="${escapeAttr(preset.value)}"><span class="shape-arrow-preview is-${sideClass} is-arrow-triangle is-arrow-size-${escapeAttr(preset.value)}"></span><span>${escapeHtml(preset.label)}</span></button>`).join("");
}



function shapeEffectsMenuHtml() {
  const glows = SHAPE_GLOW_PRESETS.map((preset) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="glow" data-shape-format-value="${escapeAttr(preset.value)}">${escapeHtml(preset.label)}</button>`).join("");
  const softEdges = SHAPE_SOFT_EDGE_PRESETS.map((preset) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="soft-edge" data-shape-format-value="${escapeAttr(preset.value)}">${escapeHtml(preset.label)}</button>`).join("");
  const reflections = SHAPE_REFLECTION_PRESETS.map((preset) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="reflection" data-shape-format-value="${escapeAttr(preset.value)}">${escapeHtml(preset.label)}</button>`).join("");
  return `<div class="shape-format-menu-title">影</div>
    <button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="shadow" data-shape-format-value="none">${escapeHtml("影なし")}</button>
    <button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="shadow" data-shape-format-value="outer1">${escapeHtml("外側: 右下")}</button>
    <button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="shadow" data-shape-format-value="outer2">${escapeHtml("外側: ぼかし")}</button>
    <div class="shape-format-menu-title">光彩</div>${glows}
    <div class="shape-format-menu-title">ぼかし</div>${softEdges}
    <div class="shape-format-menu-title">反射</div>${reflections}`;
}



function shapeTextFillMenuHtml() {
  return shapeColorSectionHtml("text-fill-color");
}



function shapeTextOutlineMenuHtml() {
  return `${shapeColorSectionHtml("text-outline-color")}
    <button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="text-outline-none">${accessKeyLabelHtml("輪郭なし(N)")}</button>`;
}



function shapeAlignMenuHtml() {
  const items = [
    ["left", "左揃え"],
    ["center", "左右中央揃え"],
    ["right", "右揃え"],
    ["top", "上揃え"],
    ["middle", "上下中央揃え"],
    ["bottom", "下揃え"],
    ["distribute-h", "左右に整列"],
    ["distribute-v", "上下に整列"],
  ];
  return items.map(([value, label]) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="align" data-shape-format-value="${escapeAttr(value)}">${escapeHtml(label)}</button>`).join("");
}



function shapeRotateMenuHtml() {
  const items = [
    ["right90", "右へ 90 度回転"],
    ["left90", "左へ 90 度回転"],
    ["flipH", "左右反転"],
    ["flipV", "上下反転"],
  ];
  return items.map(([value, label]) => `<button class="shape-format-menu-command" type="button" role="menuitem" data-shape-format-action="rotate" data-shape-format-value="${escapeAttr(value)}">${escapeHtml(label)}</button>`).join("");
}



function shapeColorSectionHtml(action) {
  const themeColors = ribbonThemeColors().slice(0, 10);
  const baseRow = themeColors.map((color) => shapeColorButtonHtml(action, color)).join("");
  const tintRows = RIBBON_COLOR_TINTS
    .map((tint) => themeColors.map((color) => shapeColorButtonHtml(action, applyTint(color, tint))).join(""))
    .join("");
  const standard = RIBBON_STANDARD_COLORS.map((color) => shapeColorButtonHtml(action, color)).join("");
  return `<div class="shape-format-menu-title">テーマ色</div>
    <div class="ribbon-color-grid ribbon-theme-color-grid">${baseRow}${tintRows}</div>
    <div class="shape-format-menu-title">標準の色</div>
    <div class="ribbon-color-grid ribbon-standard-color-grid">${standard}</div>`;
}



function shapeColorButtonHtml(action, color) {
  const normalized = normalizeColorValue(color);
  return `<button class="ribbon-color-choice" type="button" role="menuitem" title="${escapeAttr(normalized)}" style="--choice-color:${escapeAttr(normalized)}" data-shape-format-action="${escapeAttr(action)}" data-shape-format-value="${escapeAttr(normalized)}"></button>`;
}



function openShapeFormatPane() {
  const objects = selectedImageObjects();
  if (!objects.length) {
    setStatus("オブジェクトを選択してください。");
    return;
  }
  closeCellContextMenu();
  closeShapeFormatMenu();
  closeShapeFormatPane();
  state.shapeFormatPane = { previousFocus: document.activeElement };
  $("body").append(shapeFormatPaneHtml(objects));
  syncShapeFormatPane();
  syncShapeSelectionPane();
  switchRibbonTab("shape-format");
  setStatus(`${objectFormatPaneTitle(objects)}を表示しました。`);
}



function closeShapeFormatPane(options = {}) {
  document.querySelector("#shapeFormatPane")?.remove();
  const previousFocus = state.shapeFormatPane?.previousFocus;
  state.shapeFormatPane = null;
  syncShapeSelectionPane();
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) {
    previousFocus.focus?.({ preventScroll: true });
  }
}



function closeShapeFormatPaneOnEscape(event) {
  if (event.key !== "Escape" || !state.shapeFormatPane) return;
  if (event.target?.closest?.("#shapeFormatPane input, #shapeFormatPane select")) return;
  event.preventDefault();
  closeShapeFormatPane();
}



function openShapeSelectionPane() {
  if (!state.model) return;
  closeCellContextMenu();
  closeShapeFormatMenu();
  if (!state.shapeSelectionPane) {
    state.shapeSelectionPane = { previousFocus: document.activeElement };
  }
  if (!document.querySelector("#shapeSelectionPane")) {
    $("body").append(shapeSelectionPaneHtml());
  }
  syncShapeSelectionPane();
  setStatus("選択ウィンドウを表示しました。");
}



function closeShapeSelectionPane(options = {}) {
  document.querySelector("#shapeSelectionPane")?.remove();
  const previousFocus = state.shapeSelectionPane?.previousFocus;
  state.shapeSelectionPane = null;
  state.shapeSelectionPaneDrag = null;
  document.body.classList.remove("is-shape-selection-pane-dragging");
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) {
    previousFocus.focus?.({ preventScroll: true });
  }
}



function closeShapeSelectionPaneOnEscape(event) {
  if (event.key !== "Escape" || !state.shapeSelectionPane) return;
  event.preventDefault();
  closeShapeSelectionPane();
}



function shapeSelectionPaneHtml() {
  const hasFormatPane = document.querySelector("#shapeFormatPane") ? " has-format-pane" : "";
  return `<aside id="shapeSelectionPane" class="shape-selection-pane${hasFormatPane}" aria-label="選択ウィンドウ">
    <div class="shape-selection-pane-header">
      <strong>選択ウィンドウ</strong>
      <button type="button" class="shape-selection-pane-close" data-shape-selection-pane-action="close" aria-label="閉じる">×</button>
    </div>
    <div class="shape-selection-pane-toolbar">
      <button type="button" data-shape-selection-pane-action="show-all">すべて表示</button>
      <button type="button" data-shape-selection-pane-action="hide-all">すべて非表示</button>
      <button type="button" data-shape-selection-pane-action="forward" title="前面へ移動"><i data-lucide="chevron-up"></i><span>前面</span></button>
      <button type="button" data-shape-selection-pane-action="backward" title="背面へ移動"><i data-lucide="chevron-down"></i><span>背面</span></button>
    </div>
    <div class="shape-selection-pane-count"></div>
    <div class="shape-selection-pane-list" role="listbox" aria-label="オブジェクト一覧"></div>
  </aside>`;
}



function shapeSelectionPaneListHtml() {
  const sheet = activeSheet();
  const images = Array.isArray(sheet?.images) ? sheet.images : [];
  if (!images.length) {
    return '<div class="shape-selection-pane-empty">このシートにオブジェクトはありません。</div>';
  }
  const selectedIds = new Set(activeSelectedImageIds());
  return images
    .map((image, index) => ({ image, index }))
    .reverse()
    .map(({ image, index }) => {
      const selected = selectedIds.has(image.id);
      const hidden = Boolean(image.hidden);
      const label = shapeSelectionPaneObjectName(image, index);
      const renaming = state.shapeSelectionPane?.renameId === image.id;
      const rowClass = `shape-selection-pane-row${selected ? " is-selected" : ""}${hidden ? " is-hidden" : ""}`;
      const visibilityLabel = hidden ? "非表示" : "表示";
      const visibilityIcon = hidden ? "eye-off" : "eye";
      return `<div class="${rowClass}" data-image-id="${escapeAttr(image.id)}" role="option" aria-selected="${selected ? "true" : "false"}">
        <span class="shape-selection-pane-drag" data-shape-selection-pane-drag data-image-id="${escapeAttr(image.id)}" role="button" aria-label="順序変更" title="ドラッグして順序を変更">≡</span>
        <button type="button" class="shape-selection-pane-visibility" data-shape-selection-pane-action="toggle-visibility" data-image-id="${escapeAttr(image.id)}" aria-label="${visibilityLabel}" title="${visibilityLabel}" aria-pressed="${hidden ? "true" : "false"}"><i data-lucide="${visibilityIcon}"></i></button>
        ${renaming
          ? `<input class="shape-selection-pane-rename" type="text" value="${escapeAttr(label)}" data-shape-selection-pane-rename data-image-id="${escapeAttr(image.id)}" aria-label="オブジェクト名" />`
          : `<button type="button" class="shape-selection-pane-object" data-shape-selection-pane-object data-image-id="${escapeAttr(image.id)}" title="${escapeAttr(label)}">${escapeHtml(label)}</button>`}
        <button type="button" class="shape-selection-pane-rename-button" data-shape-selection-pane-action="rename" data-image-id="${escapeAttr(image.id)}">名前</button>
      </div>`;
    })
    .join("");
}



function shapeSelectionPaneObjectName(image, index) {
  const fallbackKind = image?.shape ? "図形" : "画像";
  return image?.name || image?.shape?.label || `${fallbackKind} ${index + 1}`;
}



function syncShapeSelectionPane() {
  const pane = document.querySelector("#shapeSelectionPane");
  if (!pane || !state.shapeSelectionPane || !state.model) return;
  pane.classList.toggle("has-format-pane", Boolean(document.querySelector("#shapeFormatPane")));
  const sheet = activeSheet();
  const images = Array.isArray(sheet?.images) ? sheet.images : [];
  const selectedCount = activeSelectedImageIds().filter((id) => images.some((image) => image.id === id)).length;
  const count = pane.querySelector(".shape-selection-pane-count");
  if (count) {
    const hiddenCount = images.filter((image) => image.hidden).length;
    count.textContent = `${images.length} 個のオブジェクト${hiddenCount ? ` / 非表示 ${hiddenCount}` : ""}${selectedCount ? ` / 選択 ${selectedCount}` : ""}`;
  }
  const list = pane.querySelector(".shape-selection-pane-list");
  if (list) list.innerHTML = shapeSelectionPaneListHtml();
  createIcons({ icons });
}



function handleShapeSelectionPaneAction(event) {
  const action = event.currentTarget.dataset.shapeSelectionPaneAction || "";
  const imageId = event.currentTarget.dataset.imageId || "";
  switch (action) {
    case "close":
      closeShapeSelectionPane();
      break;
    case "toggle-visibility":
      toggleShapeObjectVisibility(imageId);
      break;
    case "show-all":
      setAllShapeObjectVisibility(false);
      break;
    case "hide-all":
      setAllShapeObjectVisibility(true);
      break;
    case "forward":
      reorderSelectionPaneObjects("forward");
      break;
    case "backward":
      reorderSelectionPaneObjects("backward");
      break;
    case "rename":
      startShapeSelectionPaneRename(imageId);
      break;
    default:
      break;
  }
}



function handleShapeSelectionPaneObjectClick(event) {
  const imageId = event.currentTarget.dataset.imageId || "";
  selectShapeObjectFromPane(imageId, event);
}



function handleShapeSelectionPaneObjectDoubleClick(event) {
  const imageId = event.currentTarget.dataset.imageId || "";
  startShapeSelectionPaneRename(imageId);
}



function handleShapeSelectionPaneObjectKeydown(event) {
  const imageId = event.currentTarget.dataset.imageId || "";
  if (!imageId) return;
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "a") {
    event.preventDefault();
    event.stopPropagation();
    selectAllShapeObjectsFromPane();
    return;
  }
  if (event.key === "F2" || event.key === "Enter") {
    event.preventDefault();
    event.stopPropagation();
    startShapeSelectionPaneRename(imageId);
    return;
  }
  if (event.key === " ") {
    event.preventDefault();
    event.stopPropagation();
    toggleShapeObjectVisibility(imageId);
    return;
  }
  if (event.key === "Delete" || event.key === "Backspace") {
    event.preventDefault();
    event.stopPropagation();
    if (!activeSelectedImageIds().includes(imageId)) {
      selectShapeObjectFromPane(imageId, event);
    }
    deleteSelectedImages();
    return;
  }
  if ((event.altKey || event.metaKey) && (event.key === "ArrowUp" || event.key === "ArrowDown")) {
    event.preventDefault();
    event.stopPropagation();
    if (!activeSelectedImageIds().includes(imageId)) {
      selectShapeObjectFromPane(imageId, event);
    }
    reorderSelectionPaneObjects(event.key === "ArrowUp" ? "forward" : "backward");
    return;
  }
  if (!["ArrowUp", "ArrowDown", "Home", "End"].includes(event.key)) return;
  const visualIds = shapeSelectionPaneVisualIds();
  const currentIndex = visualIds.indexOf(imageId);
  if (currentIndex < 0) return;
  let nextIndex = currentIndex;
  if (event.key === "Home") {
    nextIndex = 0;
  } else if (event.key === "End") {
    nextIndex = visualIds.length - 1;
  } else {
    nextIndex = Math.max(0, Math.min(visualIds.length - 1, currentIndex + (event.key === "ArrowUp" ? -1 : 1)));
  }
  const nextId = visualIds[nextIndex];
  if (!nextId || nextId === imageId) return;
  event.preventDefault();
  event.stopPropagation();
  selectShapeObjectFromPane(nextId, event);
  requestAnimationFrame(() => {
    document.querySelector(`[data-shape-selection-pane-object][data-image-id="${cssEscape(nextId)}"]`)?.focus({ preventScroll: true });
  });
}



function handleShapeSelectionPaneRenameKeydown(event) {
  const input = event.currentTarget;
  if (isElementImeCompositionEvent(event, input)) return;
  if (event.key === "Enter") {
    event.preventDefault();
    commitShapeSelectionPaneRename(input.dataset.imageId || "", input.value);
  } else if (event.key === "Escape") {
    event.preventDefault();
    cancelShapeSelectionPaneRename();
  }
}



function handleShapeSelectionPaneRenameBlur(event) {
  const input = event.currentTarget;
  if (state.shapeSelectionPane?.renameId === input.dataset.imageId) {
    commitShapeSelectionPaneRename(input.dataset.imageId || "", input.value);
  }
}



function selectShapeObjectFromPane(imageId, event = {}) {
  const sheet = activeSheet();
  if (!imageId || !sheet.images?.some((image) => image.id === imageId)) return;
  const currentIds = activeSelectedImageIds();
  const imageIds = (sheet.images || []).map((image) => image.id);
  let nextIds = [imageId];
  if (event.shiftKey && currentIds.length) {
    const startIndex = imageIds.indexOf(state.selectedImage?.id || currentIds[0]);
    const endIndex = imageIds.indexOf(imageId);
    if (startIndex >= 0 && endIndex >= 0) {
      const [from, to] = startIndex < endIndex ? [startIndex, endIndex] : [endIndex, startIndex];
      nextIds = imageIds.slice(from, to + 1);
    }
  } else if (event.ctrlKey || event.metaKey) {
    const set = new Set(currentIds);
    if (set.has(imageId) && set.size > 1) {
      set.delete(imageId);
    } else {
      set.add(imageId);
    }
    nextIds = Array.from(set);
  }
  setSelectedImageIds(nextIds, imageId);
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  updateSelectionUi();
  setStatus(nextIds.length === 1 ? `${shapeSelectionPaneObjectName(selectedImageObjectById(imageId), imageIds.indexOf(imageId))}を選択しました。` : `${nextIds.length} 個のオブジェクトを選択しました。`);
}



function shapeSelectionPaneVisualIds() {
  const sheet = activeSheet();
  return [...(sheet.images || [])].reverse().map((image) => image.id);
}



function selectAllShapeObjectsFromPane() {
  const sheet = activeSheet();
  const ids = (sheet.images || []).map((image) => image.id);
  if (!ids.length) return false;
  setSelectedImageIds(ids, ids[ids.length - 1] || ids[0]);
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  updateSelectionUi();
  setStatus(`${ids.length} 個のオブジェクトを選択しました。`);
  return true;
}



function reorderSelectionPaneObjects(direction) {
  const selectedIds = activeSelectedImageIds();
  if (!selectedIds.length) {
    setStatus("順序を変更するオブジェクトを選択してください。");
    return false;
  }
  reorderSelectedObjects(direction);
  syncShapeSelectionPane();
  const focusId = state.selectedImage?.id || selectedIds[0];
  requestAnimationFrame(() => {
    document.querySelector(`[data-shape-selection-pane-object][data-image-id="${cssEscape(focusId)}"]`)?.focus({ preventScroll: true });
  });
  return true;
}



function startShapeSelectionPaneRename(imageId) {
  if (!imageId || !activeSheet().images?.some((image) => image.id === imageId)) return;
  if (!state.shapeSelectionPane) state.shapeSelectionPane = { previousFocus: document.activeElement };
  state.shapeSelectionPane.renameId = imageId;
  syncShapeSelectionPane();
  requestAnimationFrame(() => {
    const input = document.querySelector(`[data-shape-selection-pane-rename][data-image-id="${cssEscape(imageId)}"]`);
    input?.focus({ preventScroll: true });
    input?.select?.();
  });
}



function cancelShapeSelectionPaneRename() {
  if (state.shapeSelectionPane) {
    state.shapeSelectionPane.renameId = null;
    syncShapeSelectionPane();
  }
}



function commitShapeSelectionPaneRename(imageId, name) {
  const sheet = activeSheet();
  const image = sheet.images?.find((item) => item.id === imageId);
  const nextName = String(name || "").trim();
  if (!state.shapeSelectionPane || !image) return false;
  state.shapeSelectionPane.renameId = null;
  if (!nextName) {
    syncShapeSelectionPane();
    return false;
  }
  if (image.name === nextName && image.shape?.label === nextName) {
    syncShapeSelectionPane();
    return false;
  }
  if (renameObjectForSheetGroup(image, nextName)) {
    syncShapeSelectionPane();
    return true;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  image.name = nextName;
  if (image.shape) image.shape = { ...image.shape, label: nextName };
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${nextName} 名前の変更`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`オブジェクト名を ${nextName} に変更しました。`);
  return true;
}



function toggleShapeObjectVisibility(imageId) {
  const sheet = activeSheet();
  const image = sheet.images?.find((item) => item.id === imageId);
  if (!image) return;
  const restorePaneFocus = Boolean(document.activeElement?.closest?.("#shapeSelectionPane"));
  if (toggleShapeObjectVisibilityForSheetGroup(image, restorePaneFocus)) return;
  const sheetBefore = captureSheetForHistory(sheet);
  image.hidden = !image.hidden;
  const label = shapeSelectionPaneObjectName(image, sheet.images.indexOf(image));
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${label} ${image.hidden ? "非表示" : "表示"}`,
  });
  renderSheet();
  updateSelectionUi();
  if (restorePaneFocus) {
    requestAnimationFrame(() => {
      document.querySelector(`[data-shape-selection-pane-object][data-image-id="${cssEscape(imageId)}"]`)?.focus({ preventScroll: true });
    });
  }
  setStatus(`${label}を${image.hidden ? "非表示に" : "表示"}しました。`);
}



function toggleShapeObjectVisibilityForSheetGroup(image, restorePaneFocus = false) {
  const items = groupedShapeObjectTargets([image]);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const nextHidden = !image.hidden;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      if (item.images[0]) item.images[0].hidden = nextHidden;
    });
    const label = shapeSelectionPaneObjectName(image, activeSheet().images.indexOf(image));
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート ${label} ${nextHidden ? "非表示" : "表示"}`,
    });
    renderSheet();
    updateSelectionUi();
    if (restorePaneFocus) {
      requestAnimationFrame(() => {
        document.querySelector(`[data-shape-selection-pane-object][data-image-id="${cssEscape(image.id)}"]`)?.focus({ preventScroll: true });
      });
    }
    setStatus(`${items.length} 個のシートの${label}を${nextHidden ? "非表示に" : "表示"}しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクトの表示を変更できませんでした：${error.message || error}`);
  }
  return true;
}



function setAllShapeObjectVisibility(hidden) {
  const sheet = activeSheet();
  const images = sheet.images || [];
  if (!images.length || images.every((image) => Boolean(image.hidden) === hidden)) return;
  if (setAllShapeObjectVisibilityForSheetGroup(hidden)) return;
  const sheetBefore = captureSheetForHistory(sheet);
  images.forEach((image) => {
    image.hidden = hidden;
  });
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: hidden ? "すべてのオブジェクトを非表示" : "すべてのオブジェクトを表示",
  });
  renderSheet();
  updateSelectionUi();
  setStatus(hidden ? "すべてのオブジェクトを非表示にしました。" : "すべてのオブジェクトを表示しました。");
}



function setAllShapeObjectVisibilityForSheetGroup(hidden) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const targets = sheetIndexes
    .map((sheetIndex) => ({ sheetIndex, sheet: state.model?.sheets?.[sheetIndex] }))
    .filter((item) => item.sheet?.images?.length && item.sheet.images.some((image) => Boolean(image.hidden) !== hidden));
  if (targets.length <= 1) return false;
  for (const item of targets) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const sheetsBefore = captureSheetsForHistory(targets.map((item) => item.sheetIndex));
  try {
    targets.forEach((item) => {
      item.sheet.images.forEach((image) => {
        image.hidden = hidden;
      });
    });
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${targets.length} シート ${hidden ? "すべてのオブジェクトを非表示" : "すべてのオブジェクトを表示"}`,
    });
    renderSheet();
    updateSelectionUi();
    setStatus(hidden ? `${targets.length} 個のシートですべてのオブジェクトを非表示にしました。` : `${targets.length} 個のシートですべてのオブジェクトを表示しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクトの表示を変更できませんでした：${error.message || error}`);
  }
  return true;
}



function startShapeSelectionPaneDrag(event) {
  if (event.button !== 0) return;
  const imageId = event.currentTarget.dataset.imageId || "";
  if (!imageId || !activeSheet().images?.some((image) => image.id === imageId)) return;
  state.shapeSelectionPaneDrag = {
    imageId,
    pointerId: event.pointerId,
    targetId: imageId,
    placement: "before",
  };
  try {
    event.currentTarget.setPointerCapture?.(event.pointerId);
  } catch {
    // Synthetic pointer events used by tests do not always have a capturable pointer id.
  }
  document.body.classList.add("is-shape-selection-pane-dragging");
  clearShapeSelectionPaneDropMarkers();
  event.preventDefault();
  event.stopPropagation();
}



function updateShapeSelectionPaneDrag(event) {
  const job = state.shapeSelectionPaneDrag;
  if (!job) return;
  if (job.pointerId && event.pointerId && job.pointerId !== event.pointerId) return;
  const pane = document.querySelector("#shapeSelectionPane");
  if (!pane) return;
  clearShapeSelectionPaneDropMarkers();
  const rows = [...pane.querySelectorAll(".shape-selection-pane-row")];
  if (!rows.length) return;
  const row =
    document.elementFromPoint(event.clientX, event.clientY)?.closest?.(".shape-selection-pane-row") ||
    rows.reduce((best, item) => {
      const rect = item.getBoundingClientRect();
      const distance = event.clientY < rect.top ? rect.top - event.clientY : event.clientY > rect.bottom ? event.clientY - rect.bottom : 0;
      return !best || distance < best.distance ? { item, distance } : best;
    }, null)?.item;
  if (!row) return;
  const rect = row.getBoundingClientRect();
  const placement = event.clientY > rect.top + rect.height / 2 ? "after" : "before";
  job.targetId = row.dataset.imageId || job.targetId;
  job.placement = placement;
  row.classList.add(placement === "after" ? "is-drop-after" : "is-drop-before");
  event.preventDefault?.();
}



function finishShapeSelectionPaneDrag(event) {
  const job = state.shapeSelectionPaneDrag;
  if (!job) return;
  if (job.pointerId && event.pointerId && job.pointerId !== event.pointerId) return;
  state.shapeSelectionPaneDrag = null;
  document.body.classList.remove("is-shape-selection-pane-dragging");
  clearShapeSelectionPaneDropMarkers();
  reorderShapeObjectFromSelectionPane(job.imageId, job.targetId, job.placement);
  event.preventDefault?.();
}



function clearShapeSelectionPaneDropMarkers() {
  document.querySelectorAll(".shape-selection-pane-row.is-drop-before, .shape-selection-pane-row.is-drop-after").forEach((row) => {
    row.classList.remove("is-drop-before", "is-drop-after");
  });
}



function reorderShapeObjectFromSelectionPane(sourceId, targetId, placement = "before") {
  const sheet = activeSheet();
  if (!sourceId || !targetId || sourceId === targetId || !sheet.images?.length) return false;
  const imageById = new Map(sheet.images.map((image) => [image.id, image]));
  if (!imageById.has(sourceId) || !imageById.has(targetId)) return false;
  const currentVisualIds = shapeSelectionPaneVisualIds();
  const nextVisualIds = currentVisualIds.filter((id) => id !== sourceId);
  const targetIndex = nextVisualIds.indexOf(targetId);
  if (targetIndex < 0) return false;
  nextVisualIds.splice(targetIndex + (placement === "after" ? 1 : 0), 0, sourceId);
  if (nextVisualIds.join("\u0000") === currentVisualIds.join("\u0000")) return false;
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.images = nextVisualIds.reverse().map((id) => imageById.get(id)).filter(Boolean);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${shapeSelectionPaneObjectName(imageById.get(sourceId), sheet.images.indexOf(imageById.get(sourceId)))} 順序変更`,
  });
  renderSheet();
  updateSelectionUi();
  setStatus("オブジェクトの順序を変更しました。");
  return true;
}



function shapeFormatPaneHtml(objects) {
  const label = selectedObjectSummary(objects);
  const title = objectFormatPaneTitle(objects);
  const activeTab = state.shapeFormatPane?.activeTab || "fill-line";
  const tabButton = (tab, text) => {
    const active = tab === activeTab;
    return `<button type="button" class="${active ? "is-active" : ""}" role="tab" aria-selected="${active ? "true" : "false"}" data-shape-format-pane-action="tab" data-shape-format-pane-tab="${escapeAttr(tab)}">${escapeHtml(text)}</button>`;
  };
  const sectionClass = (tab) => `shape-format-pane-section${tab === activeTab ? " is-active" : ""}`;
  return `<aside id="shapeFormatPane" class="shape-format-pane" aria-label="${escapeAttr(title)}">
    <div class="shape-format-pane-header">
      <strong>${escapeHtml(title)}</strong>
      <button type="button" class="shape-format-pane-close" data-shape-format-pane-action="close" aria-label="閉じる">×</button>
    </div>
    <div class="shape-format-pane-tabs" role="tablist" aria-label="${escapeAttr(title)}">
      ${tabButton("fill-line", "塗りつぶしと線")}
      ${tabButton("effects", "効果")}
      ${tabButton("size-props", "サイズとプロパティ")}
    </div>
    <div class="shape-format-pane-body">
      <div class="shape-format-pane-selection">${escapeHtml(label)}</div>
      <section class="${sectionClass("fill-line")}" data-shape-format-pane-section="fill-line">
        <h3>塗りつぶし</h3>
        <label><span>色</span><input type="color" data-shape-format-pane-field="fillColor" /></label>
        <label><span>パターン</span><select data-shape-format-pane-field="fillPattern">${SHAPE_FILL_PATTERN_PRESETS.map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`).join("")}</select></label>
        <button type="button" class="shape-format-pane-command" data-shape-format-pane-action="fill-none">塗りつぶしなし</button>
      </section>
      <section class="${sectionClass("fill-line")}" data-shape-format-pane-section="fill-line">
        <h3>線</h3>
        <label><span>色</span><input type="color" data-shape-format-pane-field="outlineColor" /></label>
        <label><span>幅</span><input type="number" min="0.6" max="6" step="0.25" data-shape-format-pane-field="outlineWidth" /></label>
        <label><span>実線/点線</span><select data-shape-format-pane-field="lineDash">${SHAPE_LINE_DASH_PRESETS.map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`).join("")}</select></label>
        <label><span>一重/多重</span><select data-shape-format-pane-field="lineCompound">${SHAPE_LINE_COMPOUND_PRESETS.map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`).join("")}</select></label>
        <label><span>先端</span><select data-shape-format-pane-field="lineCap">${SHAPE_LINE_CAP_PRESETS.map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`).join("")}</select></label>
        <label><span>結合点</span><select data-shape-format-pane-field="lineJoin">${SHAPE_LINE_JOIN_PRESETS.map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`).join("")}</select></label>
        <label><span>矢印</span><select data-shape-format-pane-field="arrow">${SHAPE_ARROW_PRESETS.map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`).join("")}</select></label>
        <label><span>始点矢印</span><select data-shape-format-pane-field="lineStartArrowType">${shapeArrowTypeOptions()}</select></label>
        <label><span>終点矢印</span><select data-shape-format-pane-field="lineEndArrowType">${shapeArrowTypeOptions()}</select></label>
        <label><span>始点サイズ</span><select data-shape-format-pane-field="lineStartArrowSize">${shapeArrowSizeOptions()}</select></label>
        <label><span>終点サイズ</span><select data-shape-format-pane-field="lineEndArrowSize">${shapeArrowSizeOptions()}</select></label>
        <button type="button" class="shape-format-pane-command" data-shape-format-pane-action="outline-none">線なし</button>
      </section>
      <section class="${sectionClass("effects")}" data-shape-format-pane-section="effects">
        <h3>効果</h3>
        <label><span>影</span><select data-shape-format-pane-field="shadow">
          <option value="">影なし</option>
          <option value="outer1">外側: 右下</option>
          <option value="outer2">外側: ぼかし</option>
          <option value="custom">ユーザー設定</option>
        </select></label>
        <label><span>光彩</span><select data-shape-format-pane-field="glow">${SHAPE_GLOW_PRESETS.map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`).join("")}</select></label>
        <label><span>ぼかし</span><select data-shape-format-pane-field="softEdges">${SHAPE_SOFT_EDGE_PRESETS.map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`).join("")}</select></label>
        <label><span>反射</span><select data-shape-format-pane-field="reflection">${SHAPE_REFLECTION_PRESETS.map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`).join("")}</select></label>
      </section>
      <section class="${sectionClass("size-props")}" data-shape-format-pane-section="size-props">
        <h3>文字</h3>
        <label><span>フォント</span><select data-shape-format-pane-field="textFontFamily">${formatFontOptions("Calibri")}</select></label>
        <label><span>サイズ</span><select data-shape-format-pane-field="textFontSize">${shapeTextSizeOptions()}</select></label>
        <label><span>色</span><input type="color" data-shape-format-pane-field="textColor" /></label>
        <label><span>輪郭色</span><input type="color" data-shape-format-pane-field="textOutlineColor" /></label>
        <label><span>輪郭幅</span><input type="number" min="0.2" max="3" step="0.1" data-shape-format-pane-field="textOutlineWidth" /></label>
        <button type="button" class="shape-format-pane-command" data-shape-format-pane-action="text-outline-none">文字の輪郭なし</button>
        <div class="shape-format-pane-inline">
          <label><input type="checkbox" data-shape-format-pane-field="textBold" /> 太字</label>
          <label><input type="checkbox" data-shape-format-pane-field="textItalic" /> 斜体</label>
          <label><input type="checkbox" data-shape-format-pane-field="textUnderline" /> 下線</label>
        </div>
        <label><span>横位置</span><select data-shape-format-pane-field="textAlign">
          <option value="left">左揃え</option>
          <option value="center">中央揃え</option>
          <option value="right">右揃え</option>
        </select></label>
        <label><span>縦位置</span><select data-shape-format-pane-field="textVAlign">
          <option value="top">上</option>
          <option value="middle">中央</option>
          <option value="bottom">下</option>
        </select></label>
        <div class="shape-format-pane-margin-grid">
          <label><span>左余白</span><input type="number" min="0" max="80" step="1" data-shape-format-pane-field="textMarginLeft" /></label>
          <label><span>右余白</span><input type="number" min="0" max="80" step="1" data-shape-format-pane-field="textMarginRight" /></label>
          <label><span>上余白</span><input type="number" min="0" max="80" step="1" data-shape-format-pane-field="textMarginTop" /></label>
          <label><span>下余白</span><input type="number" min="0" max="80" step="1" data-shape-format-pane-field="textMarginBottom" /></label>
        </div>
        <button type="button" class="shape-format-pane-command" data-shape-format-pane-action="edit-text">テキストの編集</button>
      </section>
      <section class="${sectionClass("size-props")}" data-shape-format-pane-section="size-props">
        <h3>サイズ</h3>
        <label><span>高さ</span><input type="number" min="1" step="1" data-shape-format-pane-field="height" /></label>
        <label><span>幅</span><input type="number" min="1" step="1" data-shape-format-pane-field="width" /></label>
        <label><span>回転</span><input type="number" min="0" max="359" step="1" data-shape-format-pane-field="rotation" /></label>
      </section>
      <section class="${sectionClass("size-props")}" data-shape-format-pane-section="size-props">
        <h3>プロパティ</h3>
        <label><span>名前</span><input type="text" data-shape-format-pane-field="name" /></label>
        <button type="button" class="shape-format-pane-command" data-shape-format-pane-action="selection-pane">選択ウィンドウ...</button>
        <button type="button" class="shape-format-pane-command" data-shape-format-pane-action="object-link">リンク...</button>
      </section>
    </div>
  </aside>`;
}



function shapeTextSizeOptions(current = "11") {
  return ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36", "48", "72"]
    .map((size) => formatOption(size, size, current === size))
    .join("");
}



function shapeArrowTypeOptions(current = "none") {
  return SHAPE_ARROW_TYPE_PRESETS
    .map((preset) => formatOption(preset.value, preset.label, preset.value === current))
    .join("");
}



function shapeArrowSizeOptions(current = "medium") {
  return SHAPE_ARROW_SIZE_PRESETS
    .map((preset) => formatOption(preset.value, preset.label, preset.value === current))
    .join("");
}



function handleShapeFormatPaneAction(event) {
  const action = event.currentTarget.dataset.shapeFormatPaneAction || "";
  switch (action) {
    case "close":
      closeShapeFormatPane();
      break;
    case "tab":
      setShapeFormatPaneTab(event.currentTarget.dataset.shapeFormatPaneTab || "fill-line");
      break;
    case "fill-none":
      applySelectedShapeStyle({ fillColor: "none" }, "塗りつぶしなし");
      break;
    case "outline-none":
      applySelectedShapeStyle({ outlineColor: "none" }, "線なし");
      break;
    case "text-outline-none":
      applySelectedShapeStyle({ textOutlineColor: "none" }, "文字の輪郭なし");
      break;
    case "object-link":
      openObjectLinkDialog();
      break;
    case "selection-pane":
      openShapeSelectionPane();
      break;
    case "edit-text":
      openShapeTextEditor();
      break;
    default:
      break;
  }
}



function handleShapeFormatPaneFieldChange(event) {
  const field = event.currentTarget.dataset.shapeFormatPaneField || "";
  const value = event.currentTarget.value;
  const checked = Boolean(event.currentTarget.checked);
  if (event.currentTarget.type === "color") {
    const $control = $(event.currentTarget);
    const appliedValue = $control.data("appliedColorInput");
    $control.removeData("appliedColorInput");
    // Native color pickers emit input before an outside click can remove the pane.
    if (event.type === "input") $control.data("appliedColorInput", value);
    if (appliedValue === value) return;
  }
  switch (field) {
    case "fillColor":
      applySelectedShapeStyle({ fillColor: normalizeColorValue(value) }, "図形の塗りつぶし");
      break;
    case "outlineColor":
      applySelectedShapeStyle({ outlineColor: normalizeColorValue(value) }, "図形の枠線");
      break;
    case "fillPattern":
      applySelectedShapeStyle({
        fillPattern: value
          ? normalizeShapeFillPattern({
            preset: value,
            foreground: selectedImageObject()?.shape?.outlineColor || "#4472c4",
            background: selectedImageObject()?.shape?.fillColor || "rgba(255,255,255,0)",
          })
          : null,
      }, value ? "パターン" : "パターンなし");
      break;
    case "outlineWidth":
      applySelectedShapeStyle({ outlineWidth: normalizeShapeOutlineWidth(value) }, "線の幅");
      break;
    case "lineDash":
      applySelectedShapeStyle({ lineDash: normalizeShapeLineDash(value) }, "線の種類");
      break;
    case "lineCompound":
      applySelectedShapeStyle({ lineCompound: normalizeShapeLineCompound(value) }, "線の一重/多重");
      break;
    case "lineCap":
      applySelectedShapeStyle({ lineCap: normalizeShapeLineCap(value) }, "線の先端");
      break;
    case "lineJoin":
      applySelectedShapeStyle({ lineJoin: normalizeShapeLineJoin(value) }, "線の結合点");
      break;
    case "arrow":
      applyShapeArrowPreset(value);
      break;
    case "lineStartArrowType":
      applyShapeArrowTypePreset("start", value);
      break;
    case "lineEndArrowType":
      applyShapeArrowTypePreset("end", value);
      break;
    case "lineStartArrowSize":
      applyShapeArrowSizePreset("start", value);
      break;
    case "lineEndArrowSize":
      applyShapeArrowSizePreset("end", value);
      break;
    case "shadow":
      if (value !== "custom") applySelectedShapeStyle({ shadow: value || "" }, value ? "影" : "影なし");
      break;
    case "glow":
      applySelectedShapeStyle({ glow: normalizeShapeGlow(value) }, value ? "光彩" : "光彩なし");
      break;
    case "softEdges":
      applySelectedShapeStyle({ softEdges: normalizeShapeSoftEdges(value) }, Number(value) ? "ぼかし" : "ぼかしなし");
      break;
    case "reflection":
      applySelectedShapeStyle({ reflection: normalizeShapeReflection(value) }, value ? "反射" : "反射なし");
      break;
    case "textFontFamily":
      applySelectedShapeStyle({ textFontFamily: fontFamilyForExcel(value) }, `文字フォント：${value}`);
      break;
    case "textFontSize":
      applySelectedShapeStyle({ textFontSize: Math.max(6, Math.min(72, Number(value) || 11)) }, `文字サイズ：${value}`);
      break;
    case "textColor":
      applySelectedShapeStyle({ textColor: normalizeColorValue(value) || "#1f2937" }, "文字の塗りつぶし");
      break;
    case "textOutlineColor":
      applySelectedShapeStyle({ textOutlineColor: normalizeColorValue(value) || "#1f2937", textOutlineWidth: normalizeShapeTextOutlineWidth(selectedImageObject()?.shape?.textOutlineWidth) }, "文字の輪郭");
      break;
    case "textOutlineWidth":
      applySelectedShapeStyle({ textOutlineWidth: normalizeShapeTextOutlineWidth(value), textOutlineColor: selectedImageObject()?.shape?.textOutlineColor === "none" ? "#1f2937" : selectedImageObject()?.shape?.textOutlineColor || "#1f2937" }, "文字の輪郭幅");
      break;
    case "textBold":
      applySelectedShapeStyle({ textBold: checked }, checked ? "文字を太字" : "文字の太字解除");
      break;
    case "textItalic":
      applySelectedShapeStyle({ textItalic: checked }, checked ? "文字を斜体" : "文字の斜体解除");
      break;
    case "textUnderline":
      applySelectedShapeStyle({ textUnderline: checked ? "underline" : "" }, checked ? "文字に下線" : "文字の下線解除");
      break;
    case "textAlign":
      applySelectedShapeStyle({ textAlign: ["left", "center", "right"].includes(value) ? value : "center" }, "文字の横位置");
      break;
    case "textVAlign":
      applySelectedShapeStyle({ textVAlign: ["top", "middle", "bottom"].includes(value) ? value : "middle" }, "文字の縦位置");
      break;
    case "textMarginLeft":
    case "textMarginRight":
    case "textMarginTop":
    case "textMarginBottom":
      applySelectedShapeTextMargin(field.replace("textMargin", "").toLowerCase(), Number(value));
      break;
    case "height":
    case "width":
      applySelectedShapeSize(field, Number(value));
      break;
    case "rotation":
      applySelectedShapeRotation(Number(value));
      break;
    case "name":
      renameSelectedObjects(value);
      break;
    default:
      break;
  }
}



function setShapeFormatPaneTab(tab) {
  const normalized = ["fill-line", "effects", "size-props"].includes(tab) ? tab : "fill-line";
  if (state.shapeFormatPane) state.shapeFormatPane.activeTab = normalized;
  const pane = document.querySelector("#shapeFormatPane");
  if (!pane) return;
  pane.querySelectorAll("[data-shape-format-pane-tab]").forEach((button) => {
    const active = button.dataset.shapeFormatPaneTab === normalized;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-selected", active ? "true" : "false");
  });
  pane.querySelectorAll("[data-shape-format-pane-section]").forEach((section) => {
    section.classList.toggle("is-active", section.dataset.shapeFormatPaneSection === normalized);
  });
}



function applySelectedShapeTextMargin(side, value) {
  const normalizedSide = ["left", "right", "top", "bottom"].includes(side) ? side : "";
  if (!normalizedSide) return;
  const current = selectedImageObject()?.shape?.textMargin;
  const margin = normalizeShapeTextMargin(current, selectedImageObject()?.shape?.type === "textBox");
  margin[normalizedSide] = clampNumber(value, 0, 80, margin[normalizedSide]);
  applySelectedShapeStyle({ textMargin: margin }, "文字の余白");
}



function syncShapeFormatPane() {
  const pane = document.querySelector("#shapeFormatPane");
  if (!pane || !state.shapeFormatPane) return;
  const objects = selectedImageObjects();
  if (!objects.length) {
    closeShapeFormatPane({ restoreFocus: false });
    return;
  }
  const image = selectedImageObject() || objects[0];
  const group = selectedImageGroups(objects)[0];
  const box = group && objects.length === group.images.length ? unionImageBox(activeSheet(), group.images) : imageBoxToPixels(activeSheet(), image);
  const chartOnly = objects.length > 0 && objects.every((item) => item?.chart?.series?.length);
  const shapeObjects = objects.filter((item) => item?.shape);
  const hasShapeFields = shapeObjects.length > 0 && !chartOnly;
  const hasPictureFields = chartOnly || objects.some((item) => !item?.shape);
  const shapeImage = image?.shape ? image : shapeObjects[0];
  const shape = shapeImage?.shape || (!image?.shape ? image : {});
  const title = objectFormatPaneTitle(objects);
  const header = pane.querySelector(".shape-format-pane-header strong");
  if (header) header.textContent = title;
  pane.setAttribute("aria-label", title);
  const tablist = pane.querySelector(".shape-format-pane-tabs");
  if (tablist) tablist.setAttribute("aria-label", title);
  setShapePaneField("fillColor", shapeColorInputValue(shape.fillColor, "#dbeafe"));
  setShapePaneField("fillPattern", normalizeShapeFillPattern(shape.fillPattern)?.preset || "");
  setShapePaneField("outlineColor", shapeColorInputValue(shape.outlineColor, "#4472c4"));
  setShapePaneField("outlineWidth", String(normalizeShapeOutlineWidth(shape.outlineWidth)));
  setShapePaneField("lineDash", normalizeShapeLineDash(shape.lineDash));
  setShapePaneField("lineCompound", normalizeShapeLineCompound(shape.lineCompound));
  setShapePaneField("lineCap", normalizeShapeLineCap(shape.lineCap));
  setShapePaneField("lineJoin", normalizeShapeLineJoin(shape.lineJoin));
  setShapePaneField("arrow", shapeArrowPresetValue(shape));
  setShapePaneField("lineStartArrowType", shape.lineStartArrow ? normalizeShapeArrowType(shape.lineStartArrowType || "triangle") || "triangle" : "none");
  setShapePaneField("lineEndArrowType", shape.lineEndArrow ? normalizeShapeArrowType(shape.lineEndArrowType || "triangle") || "triangle" : "none");
  setShapePaneField("lineStartArrowSize", normalizeShapeArrowSize(shape.lineStartArrowSize));
  setShapePaneField("lineEndArrowSize", normalizeShapeArrowSize(shape.lineEndArrowSize));
  setShapePaneField("shadow", typeof shape.shadow === "object" ? "custom" : shape.shadow || "");
  setShapePaneField("glow", normalizeShapeGlow(shape.glow)?.preset || "");
  setShapePaneField("softEdges", String(normalizeShapeSoftEdges(shape.softEdges)));
  setShapePaneField("reflection", normalizeShapeReflection(shape.reflection)?.preset || "");
  setShapePaneField("textFontFamily", fontOptionFromCss(shape.textFontFamily) || "Calibri");
  setShapePaneField("textFontSize", fontSizeOptionFromCss(shape.textFontSize ? `${shape.textFontSize}pt` : "") || "11");
  setShapePaneField("textColor", shapeColorInputValue(shape.textColor, "#1f2937"));
  setShapePaneField("textOutlineColor", shapeColorInputValue(shape.textOutlineColor, "#1f2937"));
  setShapePaneField("textOutlineWidth", String(normalizeShapeTextOutlineWidth(shape.textOutlineWidth)));
  setShapePaneChecked("textBold", Boolean(shape.textBold));
  setShapePaneChecked("textItalic", Boolean(shape.textItalic));
  setShapePaneChecked("textUnderline", Boolean(underlinePresetForValue(shape.textUnderline)));
  setShapePaneField("textAlign", ["left", "center", "right"].includes(shape.textAlign) ? shape.textAlign : "center");
  setShapePaneField("textVAlign", ["top", "middle", "bottom"].includes(shape.textVAlign) ? shape.textVAlign : "middle");
  const textMargin = normalizeShapeTextMargin(shape.textMargin, shape.type === "textBox");
  setShapePaneField("textMarginLeft", String(Math.round(textMargin.left)));
  setShapePaneField("textMarginRight", String(Math.round(textMargin.right)));
  setShapePaneField("textMarginTop", String(Math.round(textMargin.top)));
  setShapePaneField("textMarginBottom", String(Math.round(textMargin.bottom)));
  setShapePaneField("height", String(Math.round(Number(box?.height || image.ext?.height || 0))));
  setShapePaneField("width", String(Math.round(Number(box?.width || image.ext?.width || 0))));
  setShapePaneField("rotation", String(Math.round(imageRotationDegrees(image))));
  setShapePaneField("name", objects.length === 1 ? image.name || image.shape?.label || (image.shape ? "図形" : "画像") : `${objects.length} 個のオブジェクト`);
  setShapePaneControlAvailability({ hasShapeFields, hasPictureFields });
  const hasLineShape = objects.some(isLineShapeObject);
  setShapePaneFieldDisabled("arrow", !hasLineShape);
  setShapePaneFieldDisabled("lineStartArrowType", !hasLineShape);
  setShapePaneFieldDisabled("lineEndArrowType", !hasLineShape);
  setShapePaneFieldDisabled("lineStartArrowSize", !hasLineShape);
  setShapePaneFieldDisabled("lineEndArrowSize", !hasLineShape);
  setShapePaneFieldDisabled("name", objects.length !== 1);
  const selection = pane.querySelector(".shape-format-pane-selection");
  if (selection) selection.textContent = selectedObjectSummary(objects);
}



function setShapePaneControlAvailability({ hasShapeFields = false, hasPictureFields = false } = {}) {
  const hasObjectVisualFields = hasShapeFields || hasPictureFields;
  const shapeOnlyFields = [
    "fillColor",
    "fillPattern",
    "lineCap",
    "lineJoin",
    "textFontFamily",
    "textFontSize",
    "textColor",
    "textOutlineColor",
    "textOutlineWidth",
    "textBold",
    "textItalic",
    "textUnderline",
    "textAlign",
    "textVAlign",
    "textMarginLeft",
    "textMarginRight",
    "textMarginTop",
    "textMarginBottom",
  ];
  const visualFields = [
    "outlineColor",
    "outlineWidth",
    "lineDash",
    "lineCompound",
    "shadow",
    "glow",
    "softEdges",
    "reflection",
  ];
  shapeOnlyFields.forEach((field) => setShapePaneFieldDisabled(field, !hasShapeFields));
  visualFields.forEach((field) => setShapePaneFieldDisabled(field, !hasObjectVisualFields));
  setShapePaneActionDisabled("fill-none", !hasShapeFields);
  setShapePaneActionDisabled("text-outline-none", !hasShapeFields);
  setShapePaneActionDisabled("edit-text", !hasShapeFields);
  setShapePaneActionDisabled("outline-none", !hasObjectVisualFields);
}



function setShapePaneField(field, value) {
  const control = document.querySelector(`[data-shape-format-pane-field="${cssEscape(field)}"]`);
  if (!control || document.activeElement === control) return;
  control.value = value;
}



function setShapePaneFieldDisabled(field, disabled) {
  const control = document.querySelector(`[data-shape-format-pane-field="${cssEscape(field)}"]`);
  if (control) control.disabled = Boolean(disabled);
}



function setShapePaneActionDisabled(action, disabled) {
  document.querySelectorAll(`[data-shape-format-pane-action="${cssEscape(action)}"]`).forEach((control) => {
    control.disabled = Boolean(disabled);
  });
}



function setShapePaneChecked(field, checked) {
  const control = document.querySelector(`[data-shape-format-pane-field="${cssEscape(field)}"]`);
  if (!control || document.activeElement === control) return;
  control.checked = Boolean(checked);
}



function shapeColorInputValue(value, fallback) {
  const color = normalizeColorValue(value);
  return color && color !== "none" ? color : fallback;
}



function shapeArrowPresetValue(shape = {}) {
  const start = Boolean(shape.lineStartArrow);
  const end = Boolean(shape.lineEndArrow);
  if (start && end) return "both";
  if (start) return "start";
  if (end) return "end";
  return "none";
}



function ribbonShapeMenuHtml() {
  const recent = (state.recentShapeTypes || DEFAULT_RECENT_SHAPE_TYPES)
    .map((type) => shapePresetForType(type))
    .filter(Boolean);
  const groups = [
    { key: "recent", label: "最近使用した図形", shapes: recent.length ? recent : DEFAULT_RECENT_SHAPE_TYPES.map((type) => shapePresetForType(type)).filter(Boolean) },
    ...RIBBON_SHAPE_GROUPS,
  ];
  const sections = groups.map((group, index) => ribbonShapeMenuSectionHtml(group, index > 0)).join("");
  return `<div id="ribbonShapeMenu" class="ribbon-shape-menu" role="menu" aria-label="図形">${sections}</div>`;
}



function ribbonShapeMenuSectionHtml(group, hasSeparator) {
  const choices = (group.shapes || []).map((shape) => ribbonShapeChoiceHtml(shape)).join("");
  return `<section class="ribbon-shape-section${hasSeparator ? " has-separator" : ""}" data-ribbon-shape-section="${escapeAttr(group.key)}">
    <div class="ribbon-shape-section-title">${escapeHtml(group.label)}</div>
    <div class="ribbon-shape-grid">${choices}</div>
  </section>`;
}



function ribbonShapeChoiceHtml(shape) {
  const preset = shapePresetForType(shape.type);
  if (!preset) return "";
  const active = preset.type === state.currentShapeType ? " is-active" : "";
  return `<button class="ribbon-shape-choice${active}" type="button" role="menuitem" title="${escapeAttr(preset.label)}" aria-label="${escapeAttr(preset.label)}" data-ribbon-shape-choice="${escapeAttr(preset.type)}">
    ${shapeSvgMarkup(preset.type, ribbonShapePreviewOptions(preset.type))}
  </button>`;
}



function ribbonShapePreviewOptions(type) {
  const base = { width: 34, height: 22 };
  switch (type) {
    case "elbowConnector":
      return {
        ...base,
        points: { x1: 0.08, y1: 0.78, x2: 0.92, y2: 0.22 },
        adjust: { bend: 0.5 },
      };
    case "curvedConnector":
      return {
        ...base,
        points: { x1: 0.08, y1: 0.74, x2: 0.92, y2: 0.26 },
        adjust: { curve: 0.08 },
      };
    default:
      return base;
  }
}



function ribbonUnderlineMenuHtml(current) {
  const currentPreset = underlinePresetForValue(current) || RIBBON_UNDERLINE_PRESETS[0];
  const choices = RIBBON_UNDERLINE_PRESETS.map((preset) => {
    const active = preset.value === currentPreset.value ? " is-active" : "";
    return `<button class="ribbon-underline-choice${active}" type="button" role="menuitem" data-ribbon-underline-choice="${escapeAttr(preset.value)}">
      ${ribbonUnderlineMarkHtml(preset)}
      <span>${accessKeyLabelHtml(preset.label)}</span>
    </button>`;
  }).join("");
  return `<div id="ribbonUnderlineMenu" class="ribbon-underline-menu" role="menu" aria-label="下線">${choices}</div>`;
}



function ribbonUnderlineMarkHtml(preset) {
  const isDouble = String(preset.value).includes("double");
  const secondLine = isDouble ? '<line x1="3" y1="13.5" x2="13" y2="13.5"></line>' : "";
  return `<span class="ribbon-underline-mark" data-ribbon-underline-mark="${escapeAttr(preset.value)}" data-ribbon-underline-letter="${escapeAttr(preset.mark)}" aria-hidden="true"><svg class="ribbon-underline-icon" viewBox="0 0 16 18" focusable="false"><text x="8" y="10.8" text-anchor="middle" text-decoration="none">${escapeHtml(preset.mark)}</text>${secondLine}<line x1="3" y1="16" x2="13" y2="16"></line></svg></span>`;
}



function openRibbonBorderMenu(event) {
  event?.preventDefault?.();
  const button = event.currentTarget;

  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeRibbonInkEraserMenu();
  closeShapeFormatMenu();

  const current = document.querySelector("[data-ribbon-border-apply]")?.dataset.ribbonBorderCurrent || "none";
  $("#ribbonBorderMenu").remove();
  $("body").append(ribbonBorderMenuHtml(current));

  const menu = document.querySelector("#ribbonBorderMenu");
  if (!menu) return;
  enhanceCommandMenuAccessKeys(menu);
  const rect = button.getBoundingClientRect();
  const top = Math.min(window.innerHeight - menu.offsetHeight - 8, rect.bottom + 2);
  const left = Math.min(window.innerWidth - menu.offsetWidth - 8, Math.max(8, rect.left - menu.offsetWidth + rect.width));
  menu.style.top = `${Math.max(8, Math.round(top))}px`;
  menu.style.left = `${Math.max(8, Math.round(left))}px`;
  state.ribbonBorderMenu = true;
  menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
}



function closeRibbonBorderMenu() {
  $("#ribbonBorderMenu").remove();
  state.ribbonBorderMenu = null;
}



function closeRibbonBorderMenuOnOutsidePointer(event) {
  if (!state.ribbonBorderMenu) return;
  if (event.target.closest?.("#ribbonBorderMenu") || event.target.closest?.(".ribbon-border-split")) return;
  closeRibbonBorderMenu();
}



function closeRibbonBorderMenuOnEscape(event) {
  if (event.key === "Escape" && state.ribbonBorderMenu) {
    if (commandMenuEscapeBlockedByDialog()) return;
    consumeCommandMenuKey(event);
    closeRibbonBorderMenu();
  }
}



function handleRibbonBorderChoice(event) {
  const preset = event.currentTarget.dataset.ribbonBorderChoice || "none";
  setRibbonBorderPreset(preset);
  closeRibbonBorderMenu();
  applyBorderPreset(preset);
}



function ribbonBorderMenuHtml(current) {
  const choices = RIBBON_BORDER_PRESETS.map((preset) => {
    const active = preset.value === current ? " is-active" : "";
    return `<button class="ribbon-border-choice${active}" type="button" role="menuitem" data-ribbon-border-choice="${escapeAttr(preset.value)}">
      <span class="ribbon-border-mark" data-ribbon-border-mark="${escapeAttr(preset.value)}" aria-hidden="true"></span>
      <span>${escapeHtml(preset.label)}</span>
    </button>`;
  }).join("");
  return `<div id="ribbonBorderMenu" class="ribbon-border-menu" role="menu" aria-label="罫線">${choices}</div>`;
}



function openRibbonColorPalette(button) {
  const style = button?.dataset?.ribbonColorMenu;
  if (!RIBBON_COLOR_STYLE[style]) return;

  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeRibbonInkEraserMenu();
  closeShapeFormatMenu();

  const rect = button.getBoundingClientRect();
  const html = ribbonColorPaletteHtml(style);
  $("#ribbonColorPalette").remove();
  $("body").append(html);

  const menu = document.querySelector("#ribbonColorPalette");
  if (!menu) return;
  enhanceCommandMenuAccessKeys(menu);
  const top = Math.min(window.innerHeight - menu.offsetHeight - 8, rect.bottom + 2);
  const left = Math.min(window.innerWidth - menu.offsetWidth - 8, Math.max(8, rect.left - menu.offsetWidth + rect.width));
  menu.style.top = `${Math.max(8, Math.round(top))}px`;
  menu.style.left = `${Math.max(8, Math.round(left))}px`;
  state.ribbonColorPalette = { style };
  if (!activeCellEditElementForFormatting()) {
    menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
  }
}



function closeRibbonColorPalette() {
  $("#ribbonColorPalette").remove();
  state.ribbonColorPalette = null;
}



function closeRibbonColorPaletteOnOutsidePointer(event) {
  if (!state.ribbonColorPalette) return;
  if (event.target.closest?.("#ribbonColorPalette") || event.target.closest?.(".ribbon-color-split") || event.target.closest?.(".ribbon-border-split")) return;
  closeRibbonColorPalette();
}



function closeRibbonColorPaletteOnEscape(event) {
  if (event.key === "Escape" && state.ribbonColorPalette) {
    if (commandMenuEscapeBlockedByDialog()) return;
    consumeCommandMenuKey(event);
    closeRibbonColorPalette();
  }
}



function handleRibbonColorChoice(event) {
  const button = event.currentTarget;
  const style = button.dataset.ribbonColorStyle;
  const color = normalizeColorValue(button.dataset.ribbonColorChoice);
  if (!style || !color) return;
  setRibbonColorValue(style, color);
  closeRibbonColorPalette();
  applyRibbonColor(style);
}



function handleRibbonColorNone(event) {
  const style = event.currentTarget.dataset.ribbonColorNone;
  const config = RIBBON_COLOR_STYLE[style];
  if (!config) return;
  closeRibbonColorPalette();
  if (style === "fontColor" && applyCellEditCssPatchToSelection({ [config.cssProperty]: null }, `${config.label}：色なし`)) return;
  applyCssToSelection({ [config.cssProperty]: null }, `${config.label}：色なし`);
}



function handleRibbonColorMore(event) {
  const style = event.currentTarget.dataset.ribbonColorMore;
  const input = ribbonColorInput(style);
  closeRibbonColorPalette();
  input?.click();
}



function ribbonColorPaletteHtml(style) {
  const themeColors = ribbonThemeColors();
  const themeBase = themeColors.slice(0, 10);
  const baseRow = themeBase.map((color) => colorChoiceButton(style, color)).join("");
  const tintRows = RIBBON_COLOR_TINTS
    .map((tint) => themeBase.map((color) => colorChoiceButton(style, applyTint(color, tint))).join(""))
    .join("");
  const standard = RIBBON_STANDARD_COLORS.map((color) => colorChoiceButton(style, color)).join("");
  const noColorLabel = "色なし(N)";

  return `
    <div id="ribbonColorPalette" class="ribbon-color-palette" role="menu" data-ribbon-color-palette="${escapeAttr(style)}">
      <div class="ribbon-color-palette-title">テーマ色</div>
      <div class="ribbon-color-grid ribbon-theme-color-grid">${baseRow}${tintRows}</div>
      <div class="ribbon-color-palette-title">標準の色</div>
      <div class="ribbon-color-grid ribbon-standard-color-grid">${standard}</div>
      <button class="ribbon-color-command" type="button" role="menuitem" data-ribbon-color-none="${escapeAttr(style)}"><span class="ribbon-color-empty-icon"></span><span>${accessKeyLabelHtml(noColorLabel)}</span></button>
      <button class="ribbon-color-command" type="button" role="menuitem" data-ribbon-color-more="${escapeAttr(style)}"><span></span><span>${accessKeyLabelHtml("その他の色(M)...")}</span></button>
    </div>
  `;
}



function colorChoiceButton(style, color) {
  const normalized = normalizeColorValue(color);
  return `<button class="ribbon-color-choice" type="button" role="menuitem" title="${escapeAttr(normalized)}" style="--choice-color:${escapeAttr(normalized)}" data-ribbon-color-style="${escapeAttr(style)}" data-ribbon-color-choice="${escapeAttr(normalized)}"></button>`;
}



function ribbonThemeColors() {
  return normalizedThemeColorPalette(state.model?.themeColors || DEFAULT_THEME_COLORS)
    .slice(0, 10)
    .filter(Boolean);
}



function normalizedThemeColorPalette(colors = DEFAULT_THEME_COLORS) {
  const source = Array.isArray(colors) ? colors : [];
  return DEFAULT_THEME_COLORS.map((fallback, index) => (
    normalizeColorValue(source[index]) || normalizeColorValue(`#${String(source[index] || "").replace("#", "")}`) || normalizeColorValue(fallback)
  )).filter(Boolean);
}



function ribbonColorInput(style) {
  return $ribbonPanels.find(`input[type="color"][data-ribbon-style="${style}"]`)[0] || null;
}



function setRibbonColorValue(style, value) {
  const color = normalizeColorValue(value);
  if (!color) return;
  const input = ribbonColorInput(style);
  if (input) input.value = color.toLowerCase();
  $ribbonPanels.find(`[data-ribbon-color-indicator="${style}"]`).css("--ribbon-color", color);
}



function setRibbonBorderPreset(preset) {
  const value = RIBBON_BORDER_PRESETS.some((item) => item.value === preset) ? preset : "none";
  const label = RIBBON_BORDER_PRESETS.find((item) => item.value === value)?.label || "";
  const main = document.querySelector("[data-ribbon-border-apply]");
  const mark = document.querySelector("[data-ribbon-border-apply] [data-ribbon-border-mark]");
  if (main) {
    main.dataset.ribbonBorderCurrent = value;
    main.setAttribute("aria-label", `${label || "罫線"}を適用`);
    main.title = label;
  }
  if (mark) {
    mark.dataset.ribbonBorderMark = value;
  }
  document.querySelectorAll("[data-ribbon-border-choice]").forEach((choice) => {
    choice.classList.toggle("is-active", choice.dataset.ribbonBorderChoice === value);
  });
}



function setRibbonUnderlinePreset(preset) {
  const item = underlinePresetForValue(preset) || RIBBON_UNDERLINE_PRESETS[0];
  const main = document.querySelector("[data-ribbon-underline-apply]");
  const mark = document.querySelector("[data-ribbon-underline-apply] [data-ribbon-underline-mark]");
  if (main) {
    main.dataset.ribbonUnderlineCurrent = item.value;
    main.setAttribute("aria-label", `${item.label.replace(/\(.+\)/, "")}を適用`);
    main.title = item.label.replace(/\(.+\)/, "");
  }
  if (mark) {
    mark.dataset.ribbonUnderlineMark = item.value;
    mark.dataset.ribbonUnderlineLetter = item.mark;
    mark.outerHTML = ribbonUnderlineMarkHtml(item);
  }
  document.querySelectorAll("[data-ribbon-underline-choice]").forEach((choice) => {
    choice.classList.toggle("is-active", choice.dataset.ribbonUnderlineChoice === item.value);
  });
}



function underlinePresetForValue(value) {
  return RIBBON_UNDERLINE_PRESETS.find((preset) => preset.value === value) || null;
}



function underlinePresetForTextDecoration(value) {
  const decoration = String(value || "").toLowerCase();
  if (!decoration.includes("underline")) return null;
  return decoration.includes("double")
    ? underlinePresetForValue("underline double")
    : underlinePresetForValue("underline");
}



function normalizeColorValue(value) {
  const raw = String(value || "").trim();
  if (/^#[0-9a-f]{6}$/i.test(raw)) return raw.toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(raw)) return `#${raw.toUpperCase()}`;
  return "";
}



function clipboardDataFromEvent(event) {
  return event?.originalEvent?.clipboardData || event?.clipboardData || null;
}



function writePlainTextToClipboardData(clipboardData, text, options = {}) {
  if (!clipboardData) return false;
  try {
    clipboardData.setData("text/plain", String(text ?? ""));
    if (options.cwsRange === true) {
      clipboardData.setData(CWS_CLIPBOARD_RANGE_MIME, CWS_CLIPBOARD_RANGE_MARKER);
    }
    if (options.cwsPayload) {
      const payloadJson = cwsClipboardPayloadJson(options.cwsPayload);
      clipboardData.setData(CWS_CLIPBOARD_DATA_MIME, payloadJson);
      clipboardData.setData("text/html", cwsClipboardHtmlForPayload(options.cwsPayload, text));
    }
    return true;
  } catch (error) {
    console.warn("Clipboard event write failed:", error);
    return false;
  }
}



function cwsClipboardPayload(kind, clipboard, text = "") {
  if (!clipboard || !kind) return null;
  return {
    format: "ChtWebSheetClipboard",
    version: 1,
    kind,
    text: String(text ?? ""),
    clipboard: cloneJsonLikeValue(clipboard),
  };
}



function cwsClipboardPayloadJson(payload) {
  return JSON.stringify(payload || {});
}



function cwsClipboardHtmlForPayload(payload, text = "") {
  const visibleHtml = payload?.kind === "range"
    ? cwsClipboardRangeHtmlTable(payload.clipboard, text)
    : `<span>${escapeHtml(text || payload?.text || "Cht WebSheet")}</span>`;
  return `${visibleHtml}${cwsClipboardHtmlPayloadComment(payload)}`;
}



function cwsClipboardHtmlPayloadComment(payload) {
  return `<!--${CWS_CLIPBOARD_HTML_MARKER}:${encodeURIComponent(cwsClipboardPayloadJson(payload))}-->`;
}



function cwsClipboardRangeHtmlTable(clipboard, text = "") {
  const rowCount = Math.max(1, Math.trunc(Number(clipboard?.rowCount) || 1));
  const colCount = Math.max(1, Math.trunc(Number(clipboard?.colCount) || 1));
  const textRows = cwsClipboardTsvRows(text);
  const entriesByKey = new Map(clipboardCellEntries(clipboard).map((cell) => [cellKey(cell.row, cell.col), cell]));
  const mergeByMaster = new Map();
  const covered = new Set();
  (clipboard?.merges || []).forEach((merge) => {
    const top = Math.max(1, Math.trunc(Number(merge.top) || 1));
    const left = Math.max(1, Math.trunc(Number(merge.left) || 1));
    const bottom = Math.min(rowCount, Math.max(top, Math.trunc(Number(merge.bottom) || top)));
    const right = Math.min(colCount, Math.max(left, Math.trunc(Number(merge.right) || left)));
    if (bottom <= top && right <= left) return;
    mergeByMaster.set(cellKey(top, left), { top, left, bottom, right });
    for (let row = top; row <= bottom; row += 1) {
      for (let col = left; col <= right; col += 1) {
        if (row !== top || col !== left) covered.add(cellKey(row, col));
      }
    }
  });
  const colHtml = Array.from({ length: colCount }, (_item, index) => {
    const width = Number(clipboard?.columns?.[index]?.width);
    return width > 0 ? `<col style="width:${roundCssNumber(width)}px">` : "<col>";
  }).join("");
  const rowHtml = Array.from({ length: rowCount }, (_rowItem, rowIndex) => {
    const row = rowIndex + 1;
    const height = Number(clipboard?.rows?.[rowIndex]?.height);
    const rowStyle = height > 0 ? ` style="height:${roundCssNumber(height)}px"` : "";
    const cells = Array.from({ length: colCount }, (_colItem, colIndex) => {
      const col = colIndex + 1;
      const key = cellKey(row, col);
      if (covered.has(key)) return "";
      const merge = mergeByMaster.get(key);
      const spanAttrs = merge
        ? `${merge.bottom > merge.top ? ` rowspan="${merge.bottom - merge.top + 1}"` : ""}${merge.right > merge.left ? ` colspan="${merge.right - merge.left + 1}"` : ""}`
        : "";
      const cell = entriesByKey.get(key);
      const value = textRows[rowIndex]?.[colIndex] ?? "";
      return `<td${spanAttrs}${cwsClipboardCellHtmlStyleAttr(cell)}>${escapeHtml(value)}</td>`;
    }).join("");
    return `<tr${rowStyle}>${cells}</tr>`;
  }).join("");
  return `<table><colgroup>${colHtml}</colgroup><tbody>${rowHtml}</tbody></table>`;
}



function cwsClipboardTsvHtmlTable(text = "") {
  const rows = cwsClipboardTsvRows(text);
  return `<table><tbody>${rows.map((row) => `<tr>${row.map((value) => `<td>${escapeHtml(value)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
}



function cwsClipboardTsvRows(text = "") {
  return String(text ?? "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => line.split("\t"));
}



function cwsClipboardCellHtmlStyleAttr(cell) {
  const declarations = cwsClipboardCellHtmlStyleDeclarations(cell);
  return declarations.length ? ` style="${escapeAttr(declarations.join(";"))}"` : "";
}



function cwsClipboardCellHtmlStyleDeclarations(cell) {
  if (!cell) return [];
  const css = cell.css || {};
  const declarations = [];
  const append = (property, value) => {
    const text = String(value ?? "").trim();
    if (text) declarations.push(`${property}:${text}`);
  };
  append("font-family", css.fontFamily);
  append("font-size", css.fontSize);
  append("color", css.color);
  append("background-color", css.backgroundColor);
  append("font-weight", css.fontWeight);
  append("font-style", css.fontStyle);
  append("text-decoration", css.textDecoration);
  append("text-align", css.textAlign);
  append("vertical-align", css.verticalAlign);
  if (css.whiteSpace) append("white-space", css.whiteSpace);
  Object.entries(cell.borders || {}).forEach(([side, border]) => {
    if (!CELL_BORDER_SIDES.includes(side)) return;
    const borderCss = cwsClipboardBorderHtmlStyle(border);
    if (borderCss) append(`border-${side}`, borderCss);
  });
  return declarations;
}



function cwsClipboardBorderHtmlStyle(border) {
  if (!border || typeof border !== "object") return "";
  const style = ["solid", "dashed", "dotted", "double"].includes(String(border.style || "").toLowerCase())
    ? String(border.style).toLowerCase()
    : "solid";
  const width = Math.max(1, Math.round(Number(border.width || border.weight) || 1));
  const color = normalizeColorValue(border.color) || "#000000";
  return `${width}px ${style} ${color}`;
}



function parseCwsClipboardPayload(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  try {
    const payload = JSON.parse(text);
    if (payload?.format !== "ChtWebSheetClipboard" || Number(payload.version) !== 1) return null;
    if (!["range", "objects"].includes(payload.kind)) return null;
    return payload;
  } catch (error) {
    return null;
  }
}



function cwsClipboardPayloadFromHtml(html) {
  const text = String(html || "");
  if (!text.includes(CWS_CLIPBOARD_HTML_MARKER)) return null;
  const commentPayload = cwsClipboardPayloadFromHtmlComment(text);
  if (commentPayload) return commentPayload;
  try {
    const doc = new DOMParser().parseFromString(text, "text/html");
    const holder = doc.querySelector(`[${CWS_CLIPBOARD_HTML_MARKER}]`);
    const payload = parseCwsClipboardPayload(holder?.textContent || "");
    if (payload) return payload;
  } catch (error) {
    // Fall through to the legacy regex parser below.
  }
  const pattern = new RegExp(`<[^>]+${escapeRegExp(CWS_CLIPBOARD_HTML_MARKER)}=["']?true["']?[^>]*>([\\s\\S]*?)<\\/[^>]+>`, "i");
  const match = text.match(pattern);
  if (!match) return null;
  return parseCwsClipboardPayload(match[1].replace(/&quot;/g, "\"").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&"));
}



function cwsClipboardPayloadFromHtmlComment(html) {
  const pattern = new RegExp(`<!--\\s*${escapeRegExp(CWS_CLIPBOARD_HTML_MARKER)}:([\\s\\S]*?)\\s*-->`, "i");
  const match = String(html || "").match(pattern);
  if (!match) return null;
  try {
    return parseCwsClipboardPayload(decodeURIComponent(match[1].trim()));
  } catch (error) {
    return null;
  }
}



function cwsClipboardPayloadFromData(clipboardData) {
  if (!clipboardData) return null;
  try {
    const direct = parseCwsClipboardPayload(clipboardData.getData(CWS_CLIPBOARD_DATA_MIME));
    if (direct) return direct;
  } catch (error) {
    // Try the HTML fallback below.
  }
  try {
    const webDirect = parseCwsClipboardPayload(clipboardData.getData(CWS_CLIPBOARD_DATA_WEB_MIME));
    if (webDirect) return webDirect;
  } catch (error) {
    // Try the HTML fallback below.
  }
  try {
    return cwsClipboardPayloadFromHtml(clipboardData.getData("text/html"));
  } catch (error) {
    return null;
  }
}



function clipboardDataText(clipboardData, type) {
  if (!clipboardData) return "";
  try {
    return clipboardData.getData(type) || "";
  } catch (error) {
    return "";
  }
}



function clipboardDataTypes(clipboardData) {
  try {
    return Array.from(clipboardData?.types || []).map((type) => String(type || "").toLowerCase());
  } catch (error) {
    return [];
  }
}



function clipboardHtmlLooksLikeCellTable(html) {
  const text = String(html || "");
  if (!text) return false;
  if (/<table[\s>]/i.test(text) && /<(td|th)[\s>]/i.test(text)) return true;
  return (
    /urn:schemas-microsoft-com:office:excel/i.test(text) ||
    /mso-(?:number-format|data-placement|style-name)/i.test(text) ||
    /<(?:x:str|x:num|x:excelworkbook)\b/i.test(text)
  );
}



function clipboardPlainTextLooksLikeCellGrid(text) {
  const normalized = String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\n$/, "");
  return normalized.includes("\t") || normalized.includes("\n");
}



function clipboardTypesLookLikeSpreadsheet(types) {
  return (types || []).some((type) => (
    type === "text/csv" ||
    type.includes("spreadsheet") ||
    type.includes("excel") ||
    type.includes("worksheet")
  ));
}



function clipboardTextPayloadLooksLikeCells(payload) {
  const text = String(payload?.text || "");
  if (!text) return false;
  return (
    clipboardHtmlLooksLikeCellTable(payload?.html) ||
    clipboardPlainTextLooksLikeCellGrid(text) ||
    clipboardTypesLookLikeSpreadsheet(payload?.types)
  );
}



function clipboardDataLooksLikeCellPaste(clipboardData, text = null) {
  if (!clipboardData) return false;
  return clipboardTextPayloadLooksLikeCells({
    text: text == null ? clipboardDataText(clipboardData, "text/plain") : text,
    html: clipboardDataText(clipboardData, "text/html"),
    types: clipboardDataTypes(clipboardData),
  });
}



function richHtmlClipboardFromData(clipboardData, text = null) {
  const html = clipboardDataText(clipboardData, "text/html");
  return richHtmlClipboardFromPayload({
    text: text == null ? clipboardDataText(clipboardData, "text/plain") : text,
    html,
    types: clipboardDataTypes(clipboardData),
  });
}



function richHtmlClipboardFromPayload(payload) {
  const html = String(payload?.html || "");
  if (!html || html.includes(CWS_CLIPBOARD_HTML_MARKER) || !clipboardHtmlLooksLikeCellTable(html)) return null;
  return excelHtmlClipboardFromHtml(html, payload?.text || "");
}



function excelHtmlClipboardFromHtml(html, fallbackText = "") {
  if (typeof DOMParser === "undefined") return null;
  let doc = null;
  try {
    doc = new DOMParser().parseFromString(String(html || ""), "text/html");
  } catch (error) {
    return null;
  }
  const table = Array.from(doc.querySelectorAll("table")).find((candidate) => candidate.querySelector("td,th"));
  if (!table) return null;

  const classStyles = excelHtmlClassStyleMap(doc);
  const columnData = excelHtmlColumnData(table, classStyles);
  const occupied = new Set();
  const textGrid = [];
  const cellEntries = [];
  const merges = [];
  const rowHeights = [];
  const columnWidths = [...columnData.widths];
  let rowCount = 0;
  let colCount = 0;

  Array.from(table.rows || []).forEach((rowElement) => {
    rowCount += 1;
    const row = rowCount;
    let col = 1;
    const rowStyle = excelHtmlStyleForElement(rowElement, classStyles);
    const rowHeight = excelHtmlElementPixelSize(rowElement, rowStyle, "height");
    if (rowHeight) rowHeights[row - 1] = rowHeight;
    textGrid[row - 1] = textGrid[row - 1] || [];

    Array.from(rowElement.cells || []).forEach((cellElement) => {
      while (occupied.has(cellKey(row, col))) col += 1;
      const rowSpan = Math.max(1, Math.trunc(Number(cellElement.rowSpan || cellElement.getAttribute("rowspan") || 1)) || 1);
      const colSpan = Math.max(1, Math.trunc(Number(cellElement.colSpan || cellElement.getAttribute("colspan") || 1)) || 1);
      const columnStyle = columnData.styles[col - 1] || {};
      const cellStyle = {
        ...columnStyle,
        ...rowStyle,
        ...excelHtmlStyleForElement(cellElement, classStyles),
      };
      const text = excelHtmlCellText(cellElement);
      const raw = excelHtmlCellRawValue(cellElement, text);
      const cellModel = excelHtmlCellModelFromStyle(cellStyle, text, raw);
      cellModel.row = row;
      cellModel.col = col;
      cellEntries.push(cellModel);
      textGrid[row - 1][col - 1] = text;

      const cellHeight = excelHtmlElementPixelSize(cellElement, cellStyle, "height");
      if (cellHeight) rowHeights[row - 1] = Math.max(rowHeights[row - 1] || 0, cellHeight);
      const cellWidth = excelHtmlElementPixelSize(cellElement, cellStyle, "width");
      if (cellWidth) {
        const widthPerColumn = cellWidth / colSpan;
        for (let offset = 0; offset < colSpan; offset += 1) {
          const targetIndex = col + offset - 1;
          columnWidths[targetIndex] = Math.max(columnWidths[targetIndex] || 0, widthPerColumn);
        }
      }

      for (let rowOffset = 0; rowOffset < rowSpan; rowOffset += 1) {
        for (let colOffset = 0; colOffset < colSpan; colOffset += 1) {
          occupied.add(cellKey(row + rowOffset, col + colOffset));
        }
      }
      if (rowSpan > 1 || colSpan > 1) {
        merges.push({
          top: row,
          left: col,
          bottom: row + rowSpan - 1,
          right: col + colSpan - 1,
        });
      }
      colCount = Math.max(colCount, col + colSpan - 1);
      col += colSpan;
    });
    colCount = Math.max(colCount, col - 1);
  });

  if (!rowCount || !colCount || !cellEntries.length) return null;
  const text = excelHtmlClipboardText(textGrid, rowCount, colCount) || String(fallbackText || "");
  return {
    text,
    mode: "copy",
    sourceSheetIndex: -1,
    sourceRange: { sheetIndex: -1, top: 1, left: 1, bottom: rowCount, right: colCount },
    rowCount,
    colCount,
    rows: Array.from({ length: rowCount }, (_item, index) => (
      rowHeights[index] ? { height: rowHeights[index], hidden: false } : null
    )),
    columns: Array.from({ length: colCount }, (_item, index) => (
      columnWidths[index] ? { width: columnWidths[index], hidden: false } : null
    )),
    merges,
    cellEntries,
    externalHtml: true,
  };
}



function excelHtmlClassStyleMap(doc) {
  const map = new Map();
  const cssText = Array.from(doc.querySelectorAll("style"))
    .map((node) => node.textContent || "")
    .join("\n")
    .replace(/<!--|-->/g, "");
  const rulePattern = /([^{}]+)\{([^{}]*)\}/g;
  let match = null;
  while ((match = rulePattern.exec(cssText))) {
    const declarations = parseCssDeclarations(match[2]);
    match[1].split(",").forEach((selector) => {
      const classPattern = /\.([_a-zA-Z][\w-]*)/g;
      let classMatch = null;
      while ((classMatch = classPattern.exec(selector))) {
        const key = classMatch[1].toLowerCase();
        map.set(key, { ...(map.get(key) || {}), ...declarations });
      }
    });
  }
  return map;
}



function excelHtmlColumnData(table, classStyles) {
  const widths = [];
  const styles = [];
  let index = 0;
  Array.from(table.querySelectorAll("col")).forEach((colElement) => {
    const span = Math.max(1, Math.trunc(Number(colElement.getAttribute("span") || 1)) || 1);
    const style = excelHtmlStyleForElement(colElement, classStyles);
    const width = excelHtmlElementPixelSize(colElement, style, "width");
    for (let offset = 0; offset < span; offset += 1) {
      styles[index] = style;
      if (width) widths[index] = width;
      index += 1;
    }
  });
  return { widths, styles };
}



function parseCssDeclarations(text) {
  const declarations = {};
  splitCssDeclarations(text).forEach((part) => {
    const index = part.indexOf(":");
    if (index <= 0) return;
    const property = part.slice(0, index).trim().toLowerCase();
    const value = part.slice(index + 1).trim();
    if (property && value) declarations[property] = value;
  });
  return declarations;
}



function splitCssDeclarations(text) {
  const parts = [];
  let current = "";
  let quote = "";
  String(text || "").replace(/\/\*[\s\S]*?\*\//g, "").split("").forEach((char) => {
    if (quote) {
      current += char;
      if (char === quote) quote = "";
      return;
    }
    if (char === "\"" || char === "'") {
      quote = char;
      current += char;
      return;
    }
    if (char === ";") {
      if (current.trim()) parts.push(current.trim());
      current = "";
      return;
    }
    current += char;
  });
  if (current.trim()) parts.push(current.trim());
  return parts;
}



function excelHtmlStyleForElement(element, classStyles) {
  const style = {};
  Array.from(element?.classList || []).forEach((className) => {
    Object.assign(style, classStyles.get(String(className || "").toLowerCase()) || {});
  });
  Object.assign(style, parseCssDeclarations(element?.getAttribute?.("style") || ""));
  const align = element?.getAttribute?.("align");
  if (align) style["text-align"] = align;
  const valign = element?.getAttribute?.("valign");
  if (valign) style["vertical-align"] = valign;
  const bgcolor = element?.getAttribute?.("bgcolor");
  if (bgcolor) style["background-color"] = bgcolor;
  return style;
}



function excelHtmlElementPixelSize(element, style, property) {
  const cssValue = style?.[property];
  const cssPixels = cssLengthToPixels(cssValue);
  if (cssPixels) return cssPixels;
  const attrValue = element?.getAttribute?.(property);
  return cssLengthToPixels(attrValue);
}



function cssLengthToPixels(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (!raw || raw === "auto") return 0;
  const match = raw.match(/-?(?:\d+(?:\.\d+)?|\.\d+)/);
  if (!match) return 0;
  const number = Number(match[0]);
  if (!Number.isFinite(number) || number <= 0) return 0;
  if (raw.includes("pt")) return Math.max(1, Math.round(number * (96 / 72)));
  if (raw.includes("in")) return Math.max(1, Math.round(number * 96));
  if (raw.includes("cm")) return Math.max(1, Math.round(number * (96 / 2.54)));
  if (raw.includes("mm")) return Math.max(1, Math.round(number * (96 / 25.4)));
  if (raw.includes("pc")) return Math.max(1, Math.round(number * 16));
  if (raw.includes("em")) return Math.max(1, Math.round(number * 16));
  return Math.max(1, Math.round(number));
}



function excelHtmlCellText(cellElement) {
  const chunks = [];
  const visit = (node) => {
    if (!node) return;
    if (node.nodeType === 3) {
      chunks.push(String(node.nodeValue || ""));
      return;
    }
    if (node.nodeType !== 1) return;
    const tag = String(node.tagName || "").toLowerCase();
    if (tag === "br") {
      chunks.push("\n");
      return;
    }
    const startsNewLine = ["div", "p"].includes(tag) && chunks.length && !chunks[chunks.length - 1].endsWith("\n");
    if (startsNewLine) chunks.push("\n");
    Array.from(node.childNodes || []).forEach(visit);
    if (["div", "p"].includes(tag) && chunks.length && !chunks[chunks.length - 1].endsWith("\n")) chunks.push("\n");
  };
  Array.from(cellElement?.childNodes || []).forEach(visit);
  return chunks.join("").replace(/\u00a0/g, " ").replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n$/, "");
}



function excelHtmlCellRawValue(cellElement, text) {
  const formula = excelHtmlAttribute(cellElement, "fmla") || excelHtmlAttribute(cellElement, "formula");
  if (formula) return String(formula).trim().startsWith("=") ? String(formula).trim() : `=${String(formula).trim()}`;
  const numeric = excelHtmlAttribute(cellElement, "num");
  if (numeric !== "") {
    const number = Number(String(numeric).replace(/,/g, ""));
    if (Number.isFinite(number)) return number;
  }
  const explicitString = excelHtmlAttribute(cellElement, "str");
  if (explicitString !== "") return explicitString;
  if (text === "") return null;
  return coerceUserInput(text, { preserveWhitespace: String(text) !== String(text).trim() });
}



function excelHtmlAttribute(element, localName) {
  const names = [localName, `x:${localName}`, `o:${localName}`];
  for (const name of names) {
    const value = element?.getAttribute?.(name);
    if (value != null) return value;
  }
  const suffix = `:${String(localName).toLowerCase()}`;
  const attr = Array.from(element?.attributes || []).find((item) => String(item.name || "").toLowerCase().endsWith(suffix));
  return attr?.value ?? "";
}



function excelHtmlCellModelFromStyle(style, text, raw) {
  const css = {};
  const fontFamily = cssFontFamilyForModel(styleValue(style, "font-family"));
  if (fontFamily) css.fontFamily = fontFamily;
  const fontSize = styleValue(style, "font-size");
  if (fontSize) css.fontSize = fontSize;
  const fontColor = htmlCssColorToHex(styleValue(style, "color"));
  if (fontColor) css.color = fontColor;
  const fillColor = htmlCssColorToHex(styleValue(style, "background-color")) || htmlCssFirstColor(styleValue(style, "background"));
  if (fillColor) css.backgroundColor = fillColor;
  const fontWeight = String(styleValue(style, "font-weight") || "").toLowerCase();
  if (fontWeight === "bold" || Number.parseInt(fontWeight, 10) >= 600) css.fontWeight = "700";
  const fontStyle = String(styleValue(style, "font-style") || "").toLowerCase();
  if (fontStyle === "italic") css.fontStyle = "italic";
  const decoration = String(styleValue(style, "text-decoration", "text-decoration-line") || "").toLowerCase();
  if (decoration.includes("underline")) css.textDecoration = decoration.includes("double") ? "underline double" : "underline";
  else if (decoration.includes("line-through")) css.textDecoration = "line-through";
  const textAlign = normalizeExcelHtmlHorizontalAlign(styleValue(style, "text-align", "mso-horizontal-align"));
  if (textAlign) css.textAlign = textAlign;
  const verticalAlign = normalizeExcelHtmlVerticalAlign(styleValue(style, "vertical-align", "mso-vertical-align"));
  if (verticalAlign) {
    css.verticalAlign = verticalAlign;
    css.alignItems = verticalAlignToFlex(verticalAlign);
  }
  if (
    String(styleValue(style, "white-space") || "").toLowerCase().includes("pre") ||
    String(styleValue(style, "mso-wrap-style") || "").toLowerCase().includes("wrap") ||
    String(text || "").includes("\n")
  ) {
    css.whiteSpace = "pre-wrap";
    css.overflowWrap = "anywhere";
  }

  const formula = typeof raw === "string" && raw.startsWith("=");
  const model = {
    raw,
    formula: formula ? raw : null,
    cached: formula ? null : raw,
    display: formula || raw == null ? "" : String(text ?? raw),
    css,
    hasFill: Boolean(css.backgroundColor),
    hasFontColor: Boolean(css.color),
    borders: excelHtmlBordersFromStyle(style),
    numFmt: excelHtmlNumberFormat(style),
  };
  if (!Object.keys(model.css).length) delete model.css;
  if (!Object.keys(model.borders).length) delete model.borders;
  if (!model.numFmt) delete model.numFmt;
  return model;
}



function styleValue(style, ...properties) {
  for (const property of properties) {
    const value = style?.[property];
    if (value != null && String(value).trim() !== "") return String(value).trim();
  }
  return "";
}



function cssFontFamilyForModel(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const first = raw.split(",")[0].replace(/^["']|["']$/g, "").trim();
  return fontFamilyForExcel(first) || raw;
}



function normalizeExcelHtmlHorizontalAlign(value) {
  const normalized = String(value || "").toLowerCase();
  if (["left", "center", "right", "justify"].includes(normalized)) return normalized;
  return "";
}



function normalizeExcelHtmlVerticalAlign(value) {
  const normalized = String(value || "").toLowerCase();
  if (normalized === "top") return "top";
  if (["middle", "center"].includes(normalized)) return "middle";
  if (normalized === "bottom") return "bottom";
  return "";
}



function excelHtmlNumberFormat(style) {
  const raw = unquoteCssValue(styleValue(style, "mso-number-format"));
  if (!raw) return "";
  const format = raw.replace(/\\(.)/g, "$1").trim();
  return /^general$/i.test(format) ? "" : format;
}



function unquoteCssValue(value) {
  const text = String(value || "").trim();
  if ((text.startsWith("\"") && text.endsWith("\"")) || (text.startsWith("'") && text.endsWith("'"))) {
    return text.slice(1, -1);
  }
  return text;
}



function htmlCssColorToHex(value) {
  const text = unquoteCssValue(value).trim().toLowerCase();
  if (!text || ["none", "transparent", "automatic"].includes(text)) return "";
  if (/^#[0-9a-f]{3}$/i.test(text)) {
    return `#${text.slice(1).split("").map((char) => char + char).join("")}`.toUpperCase();
  }
  return normalizeCellEditColor(text) || HTML_NAMED_COLORS[text] || "";
}



function htmlCssFirstColor(value) {
  const text = String(value || "").toLowerCase();
  const hex = text.match(/#[0-9a-f]{3}(?:[0-9a-f]{3})?\b/i);
  if (hex) return htmlCssColorToHex(hex[0]);
  const rgb = text.match(/rgba?\([^)]+\)/i);
  if (rgb) return htmlCssColorToHex(rgb[0]);
  const name = Object.keys(HTML_NAMED_COLORS).find((candidate) => new RegExp(`\\b${candidate}\\b`, "i").test(text));
  return name ? HTML_NAMED_COLORS[name] : "";
}



function excelHtmlBordersFromStyle(style) {
  const borders = {};
  const common = htmlCssBorderFromStyleText(styleValue(style, "border"));
  ["top", "right", "bottom", "left"].forEach((side) => {
    const sideBorder = htmlCssBorderFromStyleText(
      styleValue(style, `border-${side}`),
      styleValue(style, `border-${side}-width`),
      styleValue(style, `border-${side}-style`),
      styleValue(style, `border-${side}-color`),
    ) || common;
    if (sideBorder) borders[side] = sideBorder;
  });
  return borders;
}



function htmlCssBorderFromStyleText(borderText, widthText = "", styleText = "", colorText = "") {
  const combined = [borderText, widthText, styleText, colorText].filter(Boolean).join(" ").toLowerCase();
  if (!combined || /\b(?:none|hidden)\b/.test(combined)) return null;
  const cssStyle = combined.match(/\b(double|dashed|dotted|solid)\b/)?.[1] || "";
  if (!cssStyle) return null;
  const keywordWidth = combined.match(/\b(thin|medium|thick)\b/)?.[1] || "";
  const widthSource = combined.replace(/rgba?\([^)]+\)/gi, "");
  const explicitLength =
    widthSource.match(/(?:\d+(?:\.\d+)?|\.\d+)(?:px|pt|in|cm|mm|pc)\b/i)?.[0] ||
    widthSource.match(/\b(?:\d+(?:\.\d+)?|\.\d+)\b/i)?.[0] ||
    "";
  const width = explicitLength
    ? cssLengthToPixels(explicitLength)
    : keywordWidth === "thick"
      ? 3
      : keywordWidth === "medium"
        ? 2
        : 1;
  const color = htmlCssFirstColor(combined) || htmlCssColorToHex(colorText) || "#000000";
  return {
    width: Math.max(1, width || 1),
    style: cssStyle,
    color,
    weight: Math.max(1, width || 1) + (cssStyle === "double" ? 2 : 0),
  };
}



function excelHtmlClipboardText(textGrid, rowCount, colCount) {
  return Array.from({ length: rowCount }, (_rowItem, rowIndex) => (
    Array.from({ length: colCount }, (_colItem, colIndex) => textGrid[rowIndex]?.[colIndex] ?? "").join("\t")
  )).join("\n");
}



function clipboardDataHasCwsRangeMarker(clipboardData) {
  if (!clipboardData) return false;
  try {
    if (Array.from(clipboardData.types || []).includes(CWS_CLIPBOARD_RANGE_MIME)) return true;
    return clipboardData.getData(CWS_CLIPBOARD_RANGE_MIME) === CWS_CLIPBOARD_RANGE_MARKER;
  } catch (error) {
    return false;
  }
}



function canUseAsyncClipboardWrite() {
  return Boolean(
    navigator.clipboard?.writeText &&
      window.isSecureContext &&
      !isLocalStandaloneDocument(),
  );
}



function canUseAsyncClipboardItemWrite() {
  return Boolean(
    navigator.clipboard?.write &&
      typeof ClipboardItem === "function" &&
      window.isSecureContext &&
      !isLocalStandaloneDocument(),
  );
}



function canUseAsyncClipboardRead() {
  return Boolean(
    navigator.clipboard?.readText &&
      window.isSecureContext &&
      !isLocalStandaloneDocument(),
  );
}



function canUseAsyncClipboardItemRead() {
  return Boolean(
    navigator.clipboard?.read &&
      window.isSecureContext &&
      !isLocalStandaloneDocument(),
  );
}



function isLocalStandaloneDocument() {
  return Boolean(state.isStandalone || location.protocol === "file:" || location.origin === "null");
}



function workbookClipboardStorageSignature(model = state.model) {
  if (!model) return "";
  const sheetNames = (model.sheets || []).map((sheet) => sheet?.name || "").join("\u001f");
  return `${model.sourceName || ""}\u001e${sheetNames}`;
}



function persistedInternalClipboardPayload() {
  if (!state.internalClipboard || !state.model) return null;
  return {
    signature: workbookClipboardStorageSignature(),
    clipboard: cloneJsonLikeValue(state.internalClipboard),
  };
}



function persistInternalClipboard() {
  try {
    const payload = persistedInternalClipboardPayload();
    if (!payload) {
      sessionStorage.removeItem(INTERNAL_CLIPBOARD_STORAGE_KEY);
      return;
    }
    sessionStorage.setItem(INTERNAL_CLIPBOARD_STORAGE_KEY, JSON.stringify(payload));
  } catch (error) {
    // Session clipboard is only a paste fallback; storage failures should not affect editing.
  }
}



function clearPersistedInternalClipboard() {
  try {
    sessionStorage.removeItem(INTERNAL_CLIPBOARD_STORAGE_KEY);
  } catch (error) {
    // Ignore private-mode/local file storage restrictions.
  }
}



function normalizeClipboardRangeObject(range) {
  if (!range || typeof range !== "object") return null;
  const sheetIndex = Number.isInteger(Number(range.sheetIndex)) ? Number(range.sheetIndex) : state.activeSheetIndex;
  const top = Math.max(1, Math.trunc(Number(range.top) || Number(range.row) || 1));
  const left = Math.max(1, Math.trunc(Number(range.left) || Number(range.col) || 1));
  const bottom = Math.max(top, Math.trunc(Number(range.bottom) || top));
  const right = Math.max(left, Math.trunc(Number(range.right) || left));
  return {
    sheetIndex,
    top,
    left,
    bottom,
    right,
    fullRow: range.fullRow === true,
    fullColumn: range.fullColumn === true,
  };
}



function normalizeRestoredInternalClipboard(clipboard) {
  if (!clipboard || typeof clipboard !== "object") return null;
  const rowCount = Math.max(1, Math.trunc(Number(clipboard.rowCount) || 1));
  const colCount = Math.max(1, Math.trunc(Number(clipboard.colCount) || 1));
  return {
    text: String(clipboard.text ?? ""),
    mode: clipboard.mode === "cut" ? "cut" : "copy",
    sourceSheetIndex: Number.isInteger(Number(clipboard.sourceSheetIndex)) ? Number(clipboard.sourceSheetIndex) : state.activeSheetIndex,
    sourceRange: clipboard.sourceRange ? normalizeClipboardRangeObject(clipboard.sourceRange) : null,
    rowCount,
    colCount,
    rows: Array.isArray(clipboard.rows) ? cloneJsonLikeValue(clipboard.rows) : [],
    columns: Array.isArray(clipboard.columns) ? cloneJsonLikeValue(clipboard.columns) : [],
    merges: Array.isArray(clipboard.merges) ? cloneJsonLikeValue(clipboard.merges) : [],
    cellEntries: Array.isArray(clipboard.cellEntries) ? cloneJsonLikeValue(clipboard.cellEntries) : [],
  };
}



function normalizeRestoredObjectClipboard(clipboard) {
  if (!clipboard || typeof clipboard !== "object") return null;
  const sourceObjects = (Array.isArray(clipboard.objects) ? clipboard.objects : [])
    .filter((item) => item && typeof item === "object");
  const objects = cloneImages(sourceObjects).filter(Boolean);
  if (!objects.length) return null;
  const boxes = Array.isArray(clipboard.boxes)
    ? clipboard.boxes
        .map((box) => ({
          left: Math.round(Number(box?.left) || 0),
          top: Math.round(Number(box?.top) || 0),
          width: Math.max(1, Math.round(Number(box?.width) || 1)),
          height: Math.max(1, Math.round(Number(box?.height) || 1)),
        }))
        .filter((box) => box.width > 0 && box.height > 0)
    : [];
  return {
    type: "objects",
    mode: "copy",
    sourceSheetIndex: -1,
    sourceIds: objects.map((image, index) => String(image.id || `object-${index + 1}`)),
    primaryId: String(clipboard.primaryId || objects[0]?.id || ""),
    objects,
    boxes,
    clipboardText: String(clipboard.clipboardText || selectedObjectSummary(objects)),
    systemClipboardSynced: true,
  };
}

export {
  applyRibbonBorder,
  applyRibbonColor,
  applyRibbonUnderline,
  applySelectedShapeTextMargin,
  applyUnderlinePreset,
  calloutRectPathD,
  canUseAsyncClipboardItemRead,
  canUseAsyncClipboardItemWrite,
  canUseAsyncClipboardRead,
  canUseAsyncClipboardWrite,
  cancelShapeSelectionPaneRename,
  clearPersistedInternalClipboard,
  clearShapeSelectionPaneDropMarkers,
  clipboardDataFromEvent,
  clipboardDataHasCwsRangeMarker,
  clipboardDataLooksLikeCellPaste,
  clipboardDataText,
  clipboardDataTypes,
  clipboardHtmlLooksLikeCellTable,
  clipboardPlainTextLooksLikeCellGrid,
  clipboardTextPayloadLooksLikeCells,
  clipboardTypesLookLikeSpreadsheet,
  closeRibbonBorderMenu,
  closeRibbonBorderMenuOnEscape,
  closeRibbonBorderMenuOnOutsidePointer,
  closeRibbonColorPalette,
  closeRibbonColorPaletteOnEscape,
  closeRibbonColorPaletteOnOutsidePointer,
  closeRibbonInkEraserMenu,
  closeRibbonInkEraserMenuOnEscape,
  closeRibbonInkEraserMenuOnOutsidePointer,
  closeRibbonShapeMenu,
  closeRibbonShapeMenuOnEscape,
  closeRibbonShapeMenuOnOutsidePointer,
  closeRibbonUnderlineMenu,
  closeRibbonUnderlineMenuOnEscape,
  closeRibbonUnderlineMenuOnOutsidePointer,
  closeShapeFormatMenu,
  closeShapeFormatMenuOnEscape,
  closeShapeFormatMenuOnOutsidePointer,
  closeShapeFormatPane,
  closeShapeFormatPaneOnEscape,
  closeShapeSelectionPane,
  closeShapeSelectionPaneOnEscape,
  closeSymbolDialog,
  closeSymbolDialogOnEscape,
  cloudOutlinePoints,
  cloudPathD,
  colorChoiceButton,
  commitShapeSelectionPaneRename,
  cssFontFamilyForModel,
  cssLengthToPixels,
  cubicBezierPoint,
  cwsClipboardBorderHtmlStyle,
  cwsClipboardCellHtmlStyleAttr,
  cwsClipboardCellHtmlStyleDeclarations,
  cwsClipboardHtmlForPayload,
  cwsClipboardHtmlPayloadComment,
  cwsClipboardPayload,
  cwsClipboardPayloadFromData,
  cwsClipboardPayloadFromHtml,
  cwsClipboardPayloadFromHtmlComment,
  cwsClipboardPayloadJson,
  cwsClipboardRangeHtmlTable,
  cwsClipboardTsvHtmlTable,
  cwsClipboardTsvRows,
  defaultTextBoxDescriptor,
  excelHtmlAttribute,
  excelHtmlBordersFromStyle,
  excelHtmlCellModelFromStyle,
  excelHtmlCellRawValue,
  excelHtmlCellText,
  excelHtmlClassStyleMap,
  excelHtmlClipboardFromHtml,
  excelHtmlClipboardText,
  excelHtmlColumnData,
  excelHtmlElementPixelSize,
  excelHtmlNumberFormat,
  excelHtmlStyleForElement,
  finishShapeSelectionPaneDrag,
  fitSvgPointsString,
  handleRibbonBorderChoice,
  handleRibbonColorChoice,
  handleRibbonColorMore,
  handleRibbonColorNone,
  handleRibbonInkEraserChoice,
  handleRibbonShapeChoice,
  handleRibbonStyleChange,
  handleRibbonUnderlineChoice,
  handleShapeFormatMenuAction,
  handleShapeFormatPaneAction,
  handleShapeFormatPaneFieldChange,
  handleShapeSelectionPaneAction,
  handleShapeSelectionPaneObjectClick,
  handleShapeSelectionPaneObjectDoubleClick,
  handleShapeSelectionPaneObjectKeydown,
  handleShapeSelectionPaneRenameBlur,
  handleShapeSelectionPaneRenameKeydown,
  handleShapeStylePresetClick,
  handleSymbolChoice,
  handleSymbolDialogAction,
  htmlCssBorderFromStyleText,
  htmlCssColorToHex,
  htmlCssFirstColor,
  isLocalStandaloneDocument,
  normalizeClipboardRangeObject,
  normalizeColorValue,
  normalizeExcelHtmlHorizontalAlign,
  normalizeExcelHtmlVerticalAlign,
  normalizeRestoredInternalClipboard,
  normalizeRestoredObjectClipboard,
  normalizedThemeColorPalette,
  openLinkDialogForSelection,
  openRibbonBorderMenu,
  openRibbonColorPalette,
  openRibbonInkEraserMenu,
  openRibbonShapeMenu,
  openRibbonUnderlineMenu,
  openShapeFormatMenu,
  openShapeFormatPane,
  openShapeSelectionPane,
  openSymbolDialog,
  outlinePointsToPathD,
  parseCssDeclarations,
  parseCwsClipboardPayload,
  persistInternalClipboard,
  persistedInternalClipboardPayload,
  removeAdjacentDuplicatePoints,
  reorderSelectionPaneObjects,
  reorderShapeObjectFromSelectionPane,
  ribbonBorderMenuHtml,
  ribbonColorInput,
  ribbonColorPaletteHtml,
  ribbonInkEraserMenuHtml,
  ribbonShapeChoiceHtml,
  ribbonShapeMenuHtml,
  ribbonShapeMenuSectionHtml,
  ribbonShapePreviewOptions,
  ribbonThemeColors,
  ribbonUnderlineMarkHtml,
  ribbonUnderlineMenuHtml,
  richHtmlClipboardFromData,
  richHtmlClipboardFromPayload,
  roundPoint,
  roundSvgNumber,
  segmentIntersectionPoint,
  selectAllShapeObjectsFromPane,
  selectShapeObjectFromPane,
  setAllShapeObjectVisibility,
  setAllShapeObjectVisibilityForSheetGroup,
  setRibbonBorderPreset,
  setRibbonColorValue,
  setRibbonUnderlinePreset,
  setShapeFormatPaneTab,
  setShapePaneActionDisabled,
  setShapePaneChecked,
  setShapePaneControlAvailability,
  setShapePaneField,
  setShapePaneFieldDisabled,
  shapeAlignMenuHtml,
  shapeArrowPresetValue,
  shapeArrowSizeMenuItemsHtml,
  shapeArrowSizeOptions,
  shapeArrowTypeMenuItemsHtml,
  shapeArrowTypeOptions,
  shapeColorButtonHtml,
  shapeColorInputValue,
  shapeColorSectionHtml,
  shapeEffectsMenuHtml,
  shapeFillMenuHtml,
  shapeFormatMenuBodyHtml,
  shapeFormatMenuHtml,
  shapeFormatPaneHtml,
  shapeOutlineMenuHtml,
  shapeRotateMenuHtml,
  shapeSelectionPaneHtml,
  shapeSelectionPaneListHtml,
  shapeSelectionPaneObjectName,
  shapeSelectionPaneVisualIds,
  shapeTextFillMenuHtml,
  shapeTextOutlineMenuHtml,
  shapeTextSizeOptions,
  splitCssDeclarations,
  squaredDistance,
  startShapeSelectionPaneDrag,
  startShapeSelectionPaneRename,
  styleValue,
  svgDataUrl,
  svgPointFromLinePoint,
  svgPointsString,
  symbolDialogHtml,
  syncShapeFormatPane,
  syncShapeSelectionPane,
  toggleShapeObjectVisibility,
  toggleShapeObjectVisibilityForSheetGroup,
  underlinePresetForTextDecoration,
  underlinePresetForValue,
  uniqueWorksheetObjectId,
  unquoteCssValue,
  updateShapeSelectionPaneDrag,
  workbookClipboardStorageSignature,
  writePlainTextToClipboardData,
};
