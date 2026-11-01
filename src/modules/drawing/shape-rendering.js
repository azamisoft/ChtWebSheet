// Feature module; shared bindings are imported explicitly.
import {
  $formulaInput,
  CUSTOM_SHAPE_TYPE,
  DEFAULT_SHAPE_OUTLINE_WIDTH,
  DEFAULT_SHAPE_TYPE,
  LINE_SHAPE_TYPES,
  RIBBON_SHAPE_GROUPS,
  RIBBON_SHAPE_PRESET_MAP,
  SHAPE_FILL_PATTERN_PRESETS,
  SHAPE_GLOW_PRESETS,
  SHAPE_LINE_CAP_PRESETS,
  SHAPE_LINE_COMPOUND_PRESETS,
  SHAPE_LINE_DASH_PRESETS,
  SHAPE_LINE_JOIN_PRESETS,
  SHAPE_REFLECTION_PRESETS,
  state,
} from "../../app-core.js";
import {
  closeDataDialog,
  escapeXml,
} from "./charts-and-insertion.js";
import {
  cloudOutlinePoints,
  fitSvgPointsString,
  outlinePointsToPathD,
  removeAdjacentDuplicatePoints,
  roundPoint,
  roundSvgNumber,
  segmentIntersectionPoint,
  squaredDistance,
  svgDataUrl,
  svgPointFromLinePoint,
  svgPointsString,
  underlinePresetForValue,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  restoreGroupedSheetsAfterFailure,
  updateSelectionUi,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  selectionToolbarPoint,
  setSelectedImageIds,
  shapeAdjustmentHandles,
  shapeCustomPathHandles,
  shapeTextLayerHtml,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneJsonLikeValue,
  pushHistory,
} from "../editing/structure-and-history.js";
import {
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  columnWidthToPixels,
  escapeAttr,
  escapeHtml,
  rowHeightToPixels,
  setStatus,
} from "../ui/help-and-utilities.js";
import {
  ensureSheetSize,
} from "../ui/ribbon-and-links.js";
import {
  ensureSheetProtectionPermission,
} from "../ui/sheet-protection-and-translation.js";
import {
  rgbToHex,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  drawingCalloutPositionFromGuide,
} from "../workbook/ooxml-drawings.js";
import {
  activeSheet,
} from "../workbook/saving-and-export.js";
import {
  defaultShapeThemeEffectPatchForType,
  normalizedPageLayoutSettings,
} from "../workbook/vba-and-printing.js";


function insertEmbeddedObjectPlaceholder(options = {}) {
  const typeLabel = options.objectClassLabel || "パッケージ";
  const fileName = String(options.fileName || "").trim();
  const displayLabel = options.displayAsIcon ? (options.iconLabel || fileName || typeLabel) : (fileName || typeLabel);
  const modeLine = options.mode === "file" ? (options.linkToFile ? "リンク ファイル" : "ファイル") : "新規作成";
  const iconLine = options.displayAsIcon ? `${options.icon || "□"}  ${displayLabel}` : displayLabel;
  insertPresetTextObject("オブジェクト", {
    type: "rect",
    text: `埋め込みオブジェクト\n${iconLine}\n${modeLine}`,
    width: 220,
    height: 120,
    fillColor: "#f8fafc",
    outlineColor: "#7f8ea3",
    outlineWidth: 1,
    textColor: "#1f2937",
    textFontSize: 12,
    textFontFamily: "Yu Gothic UI, Meiryo, sans-serif",
    textBold: true,
    textAlign: "center",
    textVAlign: "middle",
    objectKind: "embeddedObject",
    objectMetadata: {
      mode: options.mode === "file" ? "file" : "create",
      objectClass: options.objectClass || "package",
      objectClassLabel: typeLabel,
      fileName,
      linkToFile: Boolean(options.linkToFile),
      displayAsIcon: Boolean(options.displayAsIcon),
      iconLabel: displayLabel,
    },
  });
}



function insertEquationObject() {
  insertPresetTextObject("数式", {
    type: "textBox",
    text: "y = mx + b",
    width: 190,
    height: 58,
    fillColor: "none",
    outlineColor: "none",
    textColor: "#111827",
    textFontSize: 24,
    textFontFamily: "Cambria Math, Cambria, serif",
    textItalic: true,
    textAlign: "center",
    textVAlign: "middle",
    textMargin: { left: 4, right: 4, top: 4, bottom: 4 },
  });
}



function openInsertIconDialog() {
  closeDataDialog({ keepState: true });
  state.dataDialog = { type: "insert-icon" };
  const overlay = document.createElement("div");
  overlay.id = "dataDialog";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = insertIconDialogHtml();
  document.body.appendChild(overlay);
  overlay.querySelector("[data-insert-icon-kind]")?.focus({ preventScroll: true });
}



function insertIconDialogHtml() {
  const choices = insertIconDefinitions().map((item) => `
    <button class="insert-gallery-card" type="button" data-data-dialog-action="insert-icon" data-insert-icon-kind="${escapeAttr(item.kind)}">
      <span class="insert-gallery-preview">${item.svg}</span>
      <span>${escapeHtml(item.name)}</span>
    </button>
  `).join("");
  return `
    <div class="excel-mini-dialog data-dialog insert-gallery-dialog" role="dialog" aria-modal="true" aria-labelledby="insertIconTitle" data-data-dialog="insert-icon">
      <div class="excel-mini-titlebar" id="insertIconTitle">アイコン</div>
      <div class="excel-mini-body">
        <div class="insert-gallery-search">
          <span>検索:</span>
          <input type="text" value="" placeholder="アイコンを検索" aria-label="アイコンを検索" data-insert-icon-search />
        </div>
        <div class="insert-gallery-grid" role="listbox" aria-label="アイコン">${choices}</div>
        <div class="insert-gallery-empty" data-insert-icon-empty hidden>一致するアイコンがありません。</div>
      </div>
      <div class="excel-mini-actions">
        <span class="excel-mini-action-spacer"></span>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>
  `;
}



function handleInsertIconSearch(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='insert-icon']");
  if (!dialog) return;
  const query = String(event.currentTarget.value || "").trim().toLowerCase();
  let visibleCount = 0;
  dialog.querySelectorAll("[data-insert-icon-kind]").forEach((button) => {
    const label = `${button.textContent || ""} ${button.dataset.insertIconKind || ""}`.toLowerCase();
    const visible = !query || label.includes(query);
    button.hidden = !visible;
    if (visible) visibleCount += 1;
  });
  const empty = dialog.querySelector("[data-insert-icon-empty]");
  if (empty) empty.hidden = visibleCount > 0;
}



function insertIconDefinitions() {
  return [
    {
      kind: "star",
      name: "星",
      svg: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M32 6l8.4 17 18.8 2.7-13.6 13.2 3.2 18.7L32 48.8 15.2 57.6l3.2-18.7L4.8 25.7 23.6 23z"/></svg>`,
    },
    {
      kind: "check",
      name: "チェック",
      svg: `<svg viewBox="0 0 64 64" aria-hidden="true"><circle cx="32" cy="32" r="26"/><path d="M19 33l8 8 18-20" fill="none" stroke="#fff" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
    },
    {
      kind: "person",
      name: "ユーザー",
      svg: `<svg viewBox="0 0 64 64" aria-hidden="true"><circle cx="32" cy="22" r="12"/><path d="M12 56c3.5-13 13-20 20-20s16.5 7 20 20z"/></svg>`,
    },
    {
      kind: "idea",
      name: "アイデア",
      svg: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M21 56h22v-7H21zM32 6c-11 0-19 8-19 18 0 7 4 12 9 16 2 2 3 4 3 6h14c0-2 1-4 3-6 5-4 9-9 9-16C51 14 43 6 32 6z"/></svg>`,
    },
    {
      kind: "gear",
      name: "設定",
      svg: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M28 5h8l3 9 9-2 4 7-7 7 7 6-4 8-9-3-3 9h-8l-3-9-9 3-4-8 7-6-7-7 4-7 9 2z"/><circle cx="32" cy="32" r="9" fill="#fff"/></svg>`,
    },
    {
      kind: "chart",
      name: "グラフ",
      svg: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M10 52h44v6H10zM14 32h8v16h-8zM28 19h8v29h-8zM42 9h8v39h-8z"/></svg>`,
    },
  ];
}



function insertIconObject(kind = "star") {
  const definition = insertIconDefinitions().find((item) => item.kind === kind) || insertIconDefinitions()[0];
  insertWorksheetImageDescriptors([
    worksheetSvgObjectDescriptor(
      "アイコン",
      worksheetIconSvg(definition.kind),
      92,
      92,
      { objectKind: "icon", label: "アイコン", iconKind: definition.kind, iconName: definition.name },
    ),
  ], { label: "アイコン" });
}



function worksheetIconSvg(kind) {
  const paths = {
    star: '<path d="M46 9l10.5 21.2 23.4 3.4-17 16.5 4 23.3L46 62.4 25.1 73.4l4-23.3-17-16.5 23.4-3.4z" fill="#4472c4"/>',
    check: '<circle cx="46" cy="46" r="34" fill="#4472c4"/><path d="M29 47l10 10 24-29" fill="none" stroke="#fff" stroke-width="8" stroke-linecap="round" stroke-linejoin="round"/>',
    person: '<circle cx="46" cy="31" r="16" fill="#4472c4"/><path d="M17 82c5-20 18-30 29-30s24 10 29 30z" fill="#4472c4"/>',
    idea: '<path d="M31 82h30v-9H31zM46 9c-15 0-26 11-26 25 0 10 6 17 13 23 3 3 4 6 4 9h18c0-3 1-6 4-9 7-6 13-13 13-23C72 20 61 9 46 9z" fill="#4472c4"/>',
    gear: '<path d="M40 8h12l4 13 13-3 6 10-10 10 10 9-6 11-13-4-4 13H40l-4-13-13 4-6-11 10-9-10-10 6-10 13 3z" fill="#4472c4"/><circle cx="46" cy="46" r="12" fill="#fff"/>',
    chart: '<path d="M15 76h62v8H15zM21 48h12v23H21zM40 31h12v40H40zM59 17h12v54H59z" fill="#4472c4"/>',
  };
  return `<svg xmlns="http://www.w3.org/2000/svg" width="92" height="92" viewBox="0 0 92 92">${paths[kind] || paths.star}</svg>`;
}



function insert3dModelPlaceholder() {
  insertWorksheetImageDescriptors([
    worksheetSvgObjectDescriptor(
      "3D モデル",
      `<svg xmlns="http://www.w3.org/2000/svg" width="180" height="150" viewBox="0 0 180 150">
        <path d="M45 44l50-28 50 28-50 28z" fill="#d9e8ff" stroke="#385d8a" stroke-width="3"/>
        <path d="M45 44v58l50 30V72z" fill="#8db4e2" stroke="#385d8a" stroke-width="3"/>
        <path d="M145 44v58l-50 30V72z" fill="#5b9bd5" stroke="#385d8a" stroke-width="3"/>
        <text x="90" y="88" text-anchor="middle" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="17" font-weight="700" fill="#17365d">3D</text>
      </svg>`,
      180,
      150,
      { objectKind: "3dModel", label: "3D モデル" },
    ),
  ], { label: "3D モデル" });
}



function insertSmartArtObject() {
  insertWorksheetImageDescriptors([
    worksheetSvgObjectDescriptor(
      "SmartArt",
      `<svg xmlns="http://www.w3.org/2000/svg" width="360" height="150" viewBox="0 0 360 150">
        <defs><marker id="smartArtArrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="8" markerHeight="8" orient="auto"><path d="M0 0l10 5-10 5z" fill="#5b6778"/></marker></defs>
        <rect x="12" y="42" width="90" height="62" rx="8" fill="#e2f0d9" stroke="#70ad47" stroke-width="3"/>
        <rect x="135" y="42" width="90" height="62" rx="8" fill="#ddebf7" stroke="#5b9bd5" stroke-width="3"/>
        <rect x="258" y="42" width="90" height="62" rx="8" fill="#fce4d6" stroke="#ed7d31" stroke-width="3"/>
        <path d="M104 73h27M227 73h27" fill="none" stroke="#5b6778" stroke-width="4" marker-end="url(#smartArtArrow)"/>
        <text x="57" y="79" text-anchor="middle" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="13" font-weight="700" fill="#385723">手順 1</text>
        <text x="180" y="79" text-anchor="middle" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="13" font-weight="700" fill="#1f4e79">手順 2</text>
        <text x="303" y="79" text-anchor="middle" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="13" font-weight="700" fill="#9e480e">手順 3</text>
      </svg>`,
      360,
      150,
      { objectKind: "smartArt", label: "SmartArt" },
    ),
  ], { label: "SmartArt" });
}



function insertScreenshotObject() {
  insertWorksheetImageDescriptors([
    worksheetSvgObjectDescriptor(
      "スクリーンショット",
      `<svg xmlns="http://www.w3.org/2000/svg" width="280" height="170" viewBox="0 0 280 170">
        <rect x="8" y="12" width="264" height="146" rx="8" fill="#f8fafc" stroke="#6b7a90" stroke-width="3"/>
        <rect x="8" y="12" width="264" height="28" rx="8" fill="#d9e2f3"/>
        <circle cx="25" cy="26" r="4" fill="#ed7d31"/><circle cx="39" cy="26" r="4" fill="#ffc000"/><circle cx="53" cy="26" r="4" fill="#70ad47"/>
        <rect x="35" y="62" width="92" height="64" fill="#dbeafe" stroke="#5b9bd5" stroke-width="2"/>
        <rect x="146" y="62" width="90" height="18" fill="#c9daf8"/><rect x="146" y="94" width="108" height="14" fill="#e7edf8"/><rect x="146" y="120" width="74" height="14" fill="#e7edf8"/>
      </svg>`,
      280,
      170,
      { objectKind: "screenshot", label: "スクリーンショット" },
    ),
  ], { label: "スクリーンショット" });
}



function openOfficeAddinsDialog(mode = "store") {
  closeDataDialog({ keepState: true });
  state.dataDialog = { type: "office-addins", mode };
  const overlay = document.createElement("div");
  overlay.id = "dataDialog";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = officeAddinsDialogHtml(mode);
  document.body.appendChild(overlay);
  overlay.querySelector(`[data-office-addin-tab="${escapeAttr(mode)}"]`)?.focus({ preventScroll: true });
}



function officeAddinsDialogHtml(mode = "store") {
  const activeMode = mode === "mine" ? "mine" : "store";
  const store = officeAddinDefinitions();
  const mine = installedOfficeAddins();
  const list = activeMode === "mine" ? mine : store;
  const cards = list.map((addin) => `
    <div class="office-addin-card" data-office-addin-id="${escapeAttr(addin.id)}">
      <div class="office-addin-icon" aria-hidden="true">${escapeHtml(addin.shortName || addin.name.slice(0, 1))}</div>
      <div class="office-addin-body">
        <strong>${escapeHtml(addin.name)}</strong>
        <span>${escapeHtml(addin.publisher)}</span>
        <p>${escapeHtml(addin.description)}</p>
      </div>
      <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="office-addin-insert" data-office-addin-id="${escapeAttr(addin.id)}">${activeMode === "mine" ? "追加" : "追加"}</button>
    </div>
  `).join("");
  return `
    <div class="excel-mini-dialog data-dialog office-addins-dialog" role="dialog" aria-modal="true" aria-labelledby="officeAddinsTitle" data-data-dialog="office-addins">
      <div class="excel-mini-titlebar" id="officeAddinsTitle">${activeMode === "mine" ? "個人用アドイン" : "Office アドイン"}</div>
      <div class="excel-mini-body">
        <div class="office-addin-tabs" role="tablist" aria-label="Office アドイン">
          <button class="${activeMode === "store" ? "is-active" : ""}" type="button" role="tab" data-data-dialog-action="office-addin-tab" data-office-addin-tab="store">ストア</button>
          <button class="${activeMode === "mine" ? "is-active" : ""}" type="button" role="tab" data-data-dialog-action="office-addin-tab" data-office-addin-tab="mine">個人用アドイン</button>
        </div>
        <div class="office-addin-list" role="list">${cards}</div>
      </div>
      <div class="excel-mini-actions">
        <span class="excel-mini-action-spacer"></span>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">閉じる</button>
      </div>
    </div>
  `;
}



function officeAddinDefinitions() {
  return [
    {
      id: "websheet-bing-maps",
      shortName: "M",
      name: "Bing Maps",
      publisher: "Microsoft",
      description: "位置情報を地図上に表示します。",
      accent: "#4472c4",
    },
    {
      id: "websheet-people-graph",
      shortName: "P",
      name: "People Graph",
      publisher: "Microsoft",
      description: "人物アイコンを使ったグラフを挿入します。",
      accent: "#70ad47",
    },
    {
      id: "websheet-mini-dashboard",
      shortName: "D",
      name: "Mini Dashboard",
      publisher: "WebSheet",
      description: "選択範囲から小さなダッシュボードを作成します。",
      accent: "#ed7d31",
    },
  ];
}



function installedOfficeAddins() {
  const installedIds = new Set(["websheet-bing-maps", "websheet-people-graph"]);
  return officeAddinDefinitions().filter((addin) => installedIds.has(addin.id));
}



function insertOfficeAddinObject(addinId) {
  const addin = officeAddinDefinitions().find((item) => item.id === addinId) || officeAddinDefinitions()[0];
  const accent = cssColorForSvg(addin.accent) || "#4472c4";
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="300" height="190" viewBox="0 0 300 190">
    <rect x="6" y="6" width="288" height="178" rx="8" fill="#ffffff" stroke="${escapeXml(accent)}" stroke-width="3"/>
    <rect x="6" y="6" width="288" height="42" rx="8" fill="${escapeXml(accent)}"/>
    <text x="24" y="34" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="18" font-weight="700" fill="#ffffff">${escapeXml(addin.name)}</text>
    <circle cx="64" cy="105" r="32" fill="${escapeXml(accent)}" opacity="0.18"/>
    <text x="64" y="116" text-anchor="middle" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="30" font-weight="700" fill="${escapeXml(accent)}">${escapeXml(addin.shortName || "A")}</text>
    <text x="118" y="94" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="14" font-weight="700" fill="#1f2937">Office アドイン</text>
    <text x="118" y="120" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="12" fill="#526173">${escapeXml(addin.publisher)}</text>
    <text x="118" y="145" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="12" fill="#526173">${escapeXml(addin.description)}</text>
  </svg>`;
  insertWorksheetImageDescriptors([
    worksheetSvgObjectDescriptor(addin.name, svg, 300, 190, {
      objectKind: "officeAddin",
      label: addin.name,
      addinId: addin.id,
      publisher: addin.publisher,
    }),
  ], { label: "アドイン" });
}



function insertSlicerObject() {
  insertPresetTextObject("スライサー", {
    type: "rect",
    text: "スライサー\n(すべて)\n項目 1\n項目 2",
    width: 180,
    height: 150,
    fillColor: "#ffffff",
    outlineColor: "#5b9bd5",
    outlineWidth: 1.2,
    textColor: "#1f2937",
    textFontSize: 12,
    textFontFamily: "Yu Gothic UI, Meiryo, sans-serif",
    textBold: false,
    textAlign: "left",
    textVAlign: "top",
    textMargin: { left: 12, right: 10, top: 10, bottom: 8 },
    objectKind: "slicer",
  });
}



function insertTimelineObject() {
  insertPresetTextObject("タイムライン", {
    type: "rect",
    text: "タイムライン\n2026\n1月   2月   3月",
    width: 300,
    height: 112,
    fillColor: "#ffffff",
    outlineColor: "#5b9bd5",
    outlineWidth: 1.2,
    textColor: "#1f2937",
    textFontSize: 12,
    textFontFamily: "Yu Gothic UI, Meiryo, sans-serif",
    textAlign: "left",
    textVAlign: "top",
    textMargin: { left: 12, right: 10, top: 10, bottom: 8 },
    objectKind: "timeline",
  });
}



function insertPresetTextObject(label, options = {}) {
  const descriptor = shapeDescriptorForType(options.type || "textBox", options);
  descriptor.name = label;
  descriptor.shape = {
    ...descriptor.shape,
    label,
    objectKind: options.objectKind || (options.type === "rect" ? "embeddedObject" : label),
    ...(options.objectMetadata ? { objectMetadata: cloneJsonLikeValue(options.objectMetadata) } : {}),
  };
  insertWorksheetImageDescriptors([descriptor], { label });
}



function worksheetSvgObjectDescriptor(name, svg, width, height, shape = {}) {
  return {
    name,
    src: svgDataUrl(svg),
    mimeType: "image/svg+xml",
    extension: "svg",
    width,
    height,
    shape: {
      ...shape,
      type: shape.type || "object",
      label: shape.label || name,
      objectKind: shape.objectKind || name,
      fillColor: "none",
      outlineColor: "none",
      outlineWidth: 0,
    },
  };
}



function openHeaderFooterDialog() {
  const sheet = activeSheet();
  if (!sheet) return;
  state.dataDialog = { type: "header-footer" };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = headerFooterDialogHtml(sheet);
  document.body.appendChild(overlay);
  overlay.querySelector("[data-header-footer-field='oddHeader.center']")?.focus({ preventScroll: true });
}



function headerFooterDialogHtml(sheet) {
  const settings = normalizedHeaderFooterSettings(sheet.headerFooter, sheet.pageLayout);
  const field = (path, label) => `
    <label class="excel-dialog-field">
      <span class="excel-dialog-caption">${escapeHtml(label)}</span>
      <input type="text" data-header-footer-field="${escapeAttr(path)}" value="${escapeAttr(headerFooterPathValue(settings, path))}" />
    </label>`;
  return `
    <div class="excel-mini-dialog data-dialog header-footer-dialog" role="dialog" aria-modal="true" aria-labelledby="headerFooterTitle" data-data-dialog="header-footer">
      <div class="excel-mini-titlebar" id="headerFooterTitle">ヘッダーとフッター</div>
      <div class="excel-mini-body data-dialog-body">
        <fieldset class="data-dialog-fieldset">
          <legend>ヘッダー</legend>
          <div class="header-footer-grid">
            ${field("oddHeader.left", "左側")}
            ${field("oddHeader.center", "中央部")}
            ${field("oddHeader.right", "右側")}
          </div>
        </fieldset>
        <fieldset class="data-dialog-fieldset">
          <legend>フッター</legend>
          <div class="header-footer-grid">
            ${field("oddFooter.left", "左側")}
            ${field("oddFooter.center", "中央部")}
            ${field("oddFooter.right", "右側")}
          </div>
        </fieldset>
        <fieldset class="data-dialog-fieldset">
          <legend>オプション</legend>
          <label class="data-dialog-check"><input type="checkbox" data-header-footer-check="differentFirst" ${settings.differentFirst ? "checked" : ""} />先頭ページのみ別指定</label>
          <label class="data-dialog-check"><input type="checkbox" data-header-footer-check="differentOddEven" ${settings.differentOddEven ? "checked" : ""} />奇数ページと偶数ページを別指定</label>
          <label class="data-dialog-check"><input type="checkbox" data-header-footer-check="scaleWithDoc" ${settings.scaleWithDoc !== false ? "checked" : ""} />ドキュメントに合わせて拡大/縮小</label>
          <label class="data-dialog-check"><input type="checkbox" data-header-footer-check="alignWithMargins" ${settings.alignWithMargins !== false ? "checked" : ""} />ページ余白に合わせて配置</label>
        </fieldset>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
}



function normalizedHeaderFooterSettings(value = {}, fallback = {}) {
  value = value || {};
  const section = (item = {}) => ({
    left: String(item.left || ""),
    center: String(item.center || ""),
    right: String(item.right || ""),
  });
  const normalized = {
    oddHeader: section(value.oddHeader),
    oddFooter: section(value.oddFooter),
    firstHeader: section(value.firstHeader),
    firstFooter: section(value.firstFooter),
    evenHeader: section(value.evenHeader),
    evenFooter: section(value.evenFooter),
    differentFirst: Boolean(value.differentFirst),
    differentOddEven: Boolean(value.differentOddEven),
    scaleWithDoc: value.scaleWithDoc !== false,
    alignWithMargins: value.alignWithMargins !== false,
  };
  const fallbackHeader = String(fallback?.headerFooter?.header ?? fallback?.header ?? "");
  const fallbackFooter = String(fallback?.headerFooter?.footer ?? fallback?.footer ?? "");
  if (!headerFooterSectionHasText(normalized.oddHeader) && fallbackHeader) normalized.oddHeader.center = fallbackHeader;
  if (!headerFooterSectionHasText(normalized.oddFooter) && fallbackFooter) normalized.oddFooter.center = fallbackFooter;
  return normalized;
}



function headerFooterSectionHasText(section) {
  return ["left", "center", "right"].some((key) => String(section?.[key] || "").trim());
}



function headerFooterSettingsFromSimpleText(headerText, footerText, current = {}) {
  const next = normalizedHeaderFooterSettings(current);
  next.oddHeader.center = String(headerText || "");
  next.oddFooter.center = String(footerText || "");
  return next;
}



function headerFooterPlainText(section = {}) {
  const values = ["left", "center", "right"].map((key) => String(section?.[key] || "").trim()).filter(Boolean);
  return values.join(" ");
}



function headerFooterPathValue(settings, path) {
  const [section, key] = String(path || "").split(".");
  return settings?.[section]?.[key] || "";
}



function applyHeaderFooterDialog(dialog) {
  const sheet = activeSheet();
  if (!sheet || !dialog) return;
  const before = captureSheetForHistory(sheet);
  const next = normalizedHeaderFooterSettings(sheet.headerFooter, sheet.pageLayout);
  dialog.querySelectorAll("[data-header-footer-field]").forEach((input) => {
    const [section, key] = String(input.dataset.headerFooterField || "").split(".");
    if (!next[section] || !key) return;
    next[section][key] = input.value || "";
  });
  next.differentFirst = Boolean(dialog.querySelector("[data-header-footer-check='differentFirst']")?.checked);
  next.differentOddEven = Boolean(dialog.querySelector("[data-header-footer-check='differentOddEven']")?.checked);
  next.scaleWithDoc = Boolean(dialog.querySelector("[data-header-footer-check='scaleWithDoc']")?.checked);
  next.alignWithMargins = Boolean(dialog.querySelector("[data-header-footer-check='alignWithMargins']")?.checked);
  sheet.headerFooter = next;
  sheet.pageLayout = normalizedPageLayoutSettings(sheet.pageLayout, sheet);
  sheet.pageLayout.headerFooter = {
    header: headerFooterPlainText(next.oddHeader),
    footer: headerFooterPlainText(next.oddFooter),
  };
  closeDataDialog();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore: before,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: "ヘッダーとフッター",
  });
  scheduleWorkbookPersistence();
  setStatus("ヘッダーとフッターを設定しました。");
}



function insertWorksheetImageDescriptors(descriptors, options = {}) {
  const cleanDescriptors = (descriptors || []).filter((descriptor) => descriptor?.src);
  if (!cleanDescriptors.length) return;
  const sheet = activeSheet();
  if (!options.skipSheetGroup && insertWorksheetImageDescriptorsToSheetGroup(cleanDescriptors, options)) return;
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを挿入できません。")) return;
  const anchor = options.anchor || insertionPointForObjects();
  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  const inserted = addWorksheetImageDescriptorsToSheet(sheet, cleanDescriptors, anchor, options);
  if (!inserted.length) return;

  const first = inserted[0];
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  state.selectionDrag = null;
  state.formulaEdit = null;
  setSelectedImageIds(inserted.map((image) => image.id), first.id);
  $formulaInput.prop("disabled", true).val("");
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore,
    selectionAfter: null,
    selectionBeforeRanges,
    selectionAfterRanges: [],
    label: `${options.label || "画像"}の挿入`,
  });
  renderSheet();
  updateSelectionUi();
  setStatus(`${inserted.length} 個の${options.label || "画像"}を挿入しました。`);
}



function insertWorksheetImageDescriptorsToSheetGroup(descriptors, options = {}) {
  const cleanDescriptors = (descriptors || []).filter((descriptor) => descriptor?.src);
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !cleanDescriptors.length) return false;
  const sourceAnchor = options.anchor || insertionPointForObjects();
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを挿入できません。")) return true;
    items.push({
      sheetIndex,
      sheet,
      anchor: {
        ...sourceAnchor,
        sheetIndex,
      },
    });
  }
  if (!items.length) return true;

  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  const label = options.label || "画像";
  try {
    let activeInserted = [];
    items.forEach((item) => {
      const inserted = addWorksheetImageDescriptorsToSheet(item.sheet, cleanDescriptors, item.anchor, options);
      if (item.sheetIndex === state.activeSheetIndex) activeInserted = inserted;
    });
    if (!activeInserted.length) throw new Error("No object was inserted on the active sheet.");
    state.selected = null;
    state.selectionRange = null;
    state.selectionRanges = [];
    state.selectionAnchor = null;
    state.selectionStart = null;
    state.selectionDrag = null;
    state.formulaEdit = null;
    setSelectedImageIds(activeInserted.map((image) => image.id), activeInserted[0].id);
    $formulaInput.prop("disabled", true).val("");
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore,
      selectionAfter: null,
      selectionBeforeRanges,
      selectionAfterRanges: [],
      label: `${sheetIndexes.length} シート${label}の挿入`,
    });
    renderSheet();
    updateSelectionUi();
    setStatus(`${sheetIndexes.length} 個のシートに${activeInserted.length} 個の${label}を挿入しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`${label}を挿入できませんでした：${error.message || error}`);
  }
  return true;
}



function addWorksheetImageDescriptorsToSheet(sheet, descriptors, anchor, options = {}) {
  const cleanDescriptors = (descriptors || []).filter((descriptor) => descriptor?.src);
  if (!sheet || !cleanDescriptors.length || !anchor) return [];
  const inserted = [];
  let row = anchor.row;
  cleanDescriptors.forEach((descriptor) => {
    const size = descriptor.preserveDisplaySize
      ? objectDisplaySize(descriptor.width, descriptor.height)
      : fitObjectSize(descriptor.width, descriptor.height);
    ensureSheetSize(sheet, row, anchor.col);
    const end = endingCellForPixelSize(sheet, row, anchor.col, size.width, size.height);
    ensureSheetSize(sheet, end.row, end.col);
    const image = {
      id: uniqueWorksheetObjectId("image"),
      name: descriptor.name || options.label || "画像",
      src: descriptor.src,
      mimeType: descriptor.mimeType || "image/png",
      extension: descriptor.extension || "png",
      editAs: "oneCell",
      tl: {
        col: anchor.col - 1,
        row: row - 1,
        nativeCol: anchor.col - 1,
        nativeRow: row - 1,
        nativeColOff: 0,
        nativeRowOff: 0,
      },
      br: null,
      ext: { width: size.width, height: size.height },
      hyperlink: "",
    };
    if (descriptor.chart) image.chart = cloneJsonLikeValue(descriptor.chart);
    if (descriptor.shape) image.shape = cloneJsonLikeValue(descriptor.shape);
    sheet.images = Array.isArray(sheet.images) ? sheet.images : [];
    sheet.images.push(image);
    inserted.push(image);
    row = end.row + 1;
  });
  return inserted;
}



function insertionPointForObjects() {
  const point = selectionToolbarPoint();
  if (point?.sheetIndex === state.activeSheetIndex) return point;
  const range = activeSelectionRange();
  if (range) return { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
  return { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
}



function fitObjectSize(width, height) {
  const rawWidth = Math.max(24, Number(width) || 220);
  const rawHeight = Math.max(18, Number(height) || 140);
  const scale = Math.min(1, 420 / rawWidth, 280 / rawHeight);
  return {
    width: Math.round(rawWidth * scale),
    height: Math.round(rawHeight * scale),
  };
}



function objectDisplaySize(width, height) {
  return {
    width: Math.max(1, Math.round(Number(width) || 220)),
    height: Math.max(1, Math.round(Number(height) || 140)),
  };
}



function endingCellForPixelSize(sheet, startRow, startCol, width, height) {
  return {
    row: endingIndexForPixelSpan(sheet.rows, startRow, sheet.defaultRowHeight || rowHeightToPixels(15), height),
    col: endingIndexForPixelSpan(sheet.columns, startCol, sheet.defaultColumnWidth || columnWidthToPixels(8.43), width),
  };
}



function endingIndexForPixelSpan(items, startIndex, defaultSize, pixels) {
  let total = 0;
  let index = startIndex;
  const maxIndex = startIndex + 200;
  while (total < pixels && index <= maxIndex) {
    const item = items[index - 1];
    total += item?.hidden ? 0 : item?.width || item?.height || defaultSize;
    index += 1;
  }
  return Math.max(startIndex, index - 1);
}



function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error("画像を読み込めません。"));
    reader.readAsDataURL(file);
  });
}



function imageDimensionsFromSource(src) {
  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => resolve({ width: image.naturalWidth || 220, height: image.naturalHeight || 140 });
    image.onerror = () => resolve({ width: 220, height: 140 });
    image.src = src;
  });
}



function fileExtension(name) {
  const match = String(name || "").match(/\.([^.]+)$/);
  return match ? match[1].toLowerCase() : "png";
}



function fileExtensionFromImageMimeType(mimeType, fallback = "png") {
  const normalized = String(mimeType || "").toLowerCase();
  if (normalized.includes("jpeg") || normalized.includes("jpg")) return "jpeg";
  if (normalized.includes("png")) return "png";
  if (normalized.includes("gif")) return "gif";
  if (normalized.includes("webp")) return "webp";
  if (normalized.includes("bmp")) return "bmp";
  if (normalized.includes("svg")) return "svg";
  if (normalized.includes("icon") || normalized.includes("x-icon")) return "ico";
  return String(fallback || "png").replace(/^\./, "").toLowerCase() || "png";
}



function defaultShapeDescriptor() {
  return shapeDescriptorForType(DEFAULT_SHAPE_TYPE);
}



function shapeDescriptorForType(type, style = {}) {
  const preset = shapePresetForType(type) || shapePresetForType(DEFAULT_SHAPE_TYPE);
  const isLine = isLineShapeType(preset.type);
  const isTextBox = preset.type === "textBox";
  const themeEffectPatch = defaultShapeThemeEffectPatchForType(preset.type, style);
  const customPath = preset.type === CUSTOM_SHAPE_TYPE ? normalizeShapeCustomPath(style.customPath) : null;
  const width = Math.max(1, Number(style.width) || (isLine ? 180 : isTextBox ? 220 : 150));
  const height = Math.max(1, Number(style.height) || (isLine ? 70 : isTextBox ? 64 : 100));
  const fillColor = isLine || isTextBox ? style.fillColor ?? "none" : style.fillColor ?? "#eaf3ff";
  const fillGradient = isLine ? null : normalizeShapeGradient(style.fillGradient);
  const fillPattern = isLine ? null : normalizeShapeFillPattern(style.fillPattern);
  const outlineColor = style.outlineColor ?? (isTextBox ? "none" : "#4472c4");
  const outlineWidth = normalizeShapeOutlineWidth(style.outlineWidth);
  const lineDash = normalizeShapeLineDash(style.lineDash);
  const lineCompound = normalizeShapeLineCompound(style.lineCompound);
  const lineCap = normalizeShapeLineCap(style.lineCap);
  const lineJoin = normalizeShapeLineJoin(style.lineJoin);
  const rotation = normalizeRotationDegrees(style.rotation);
  const flipH = Boolean(style.flipH);
  const flipV = Boolean(style.flipV);
  const points = isLine ? normalizeShapeLinePoints(style.points) : null;
  const lineStartArrow = isLine ? Boolean(style.lineStartArrow ?? preset.type === "doubleArrow") : false;
  const lineEndArrow = isLine ? Boolean(style.lineEndArrow ?? ["arrow", "doubleArrow", "elbowConnector", "curvedConnector"].includes(preset.type)) : false;
  const lineStartArrowType = lineStartArrow ? normalizeShapeArrowType(style.lineStartArrowType || "triangle") : "";
  const lineEndArrowType = lineEndArrow ? normalizeShapeArrowType(style.lineEndArrowType || "triangle") : "";
  const lineStartArrowSize = lineStartArrow ? normalizeShapeArrowSize(style.lineStartArrowSize) : "";
  const lineEndArrowSize = lineEndArrow ? normalizeShapeArrowSize(style.lineEndArrowSize) : "";
  const connections = isLine ? normalizeLineShapeConnections(style.connections) : null;
  const adjust = normalizeShapeAdjustments(preset.type, style.adjust);
  const textRuns = isLine ? [] : normalizeShapeTextRuns(style.textRuns);
  const text = isLine ? "" : String(style.text ?? shapePlainTextFromRuns(textRuns) ?? "");
  const textColor = style.textColor || "#1f2937";
  const textFontSize = clampNumber(style.textFontSize, 6, 72, 11);
  const textAlign = ["left", "center", "right"].includes(style.textAlign) ? style.textAlign : isTextBox ? "left" : "center";
  const textVAlign = ["top", "middle", "bottom"].includes(style.textVAlign) ? style.textVAlign : isTextBox ? "top" : "middle";
  const textBold = Boolean(style.textBold);
  const textItalic = Boolean(style.textItalic);
  const textUnderline = underlinePresetForValue(style.textUnderline)?.value || "";
  const textFontFamily = style.textFontFamily || "Yu Gothic UI, Meiryo, sans-serif";
  const textMargin = normalizeShapeTextMargin(style.textMargin, isTextBox);
  const textOutlineColor = style.textOutlineColor === "none" ? "none" : cssColorForSvg(style.textOutlineColor) || "none";
  const textOutlineWidth = normalizeShapeTextOutlineWidth(style.textOutlineWidth);
  const includeTextInSvg = style.includeTextInSvg !== false;
  const shadow = normalizeShapeShadow(Object.prototype.hasOwnProperty.call(style, "shadow") ? style.shadow : themeEffectPatch.shadow);
  const glow = normalizeShapeGlow(Object.prototype.hasOwnProperty.call(style, "glow") ? style.glow : themeEffectPatch.glow);
  const softEdges = normalizeShapeSoftEdges(Object.prototype.hasOwnProperty.call(style, "softEdges") ? style.softEdges : themeEffectPatch.softEdges);
  const reflection = normalizeShapeReflection(Object.prototype.hasOwnProperty.call(style, "reflection") ? style.reflection : themeEffectPatch.reflection);
  const svg = shapeSvgMarkup(preset.type, { width, height, fillColor, fillGradient, fillPattern, outlineColor, outlineWidth, lineDash, lineCompound, lineCap, lineJoin, points, lineStartArrow, lineEndArrow, lineStartArrowType, lineEndArrowType, lineStartArrowSize, lineEndArrowSize, adjust, customPath, shadow, glow, softEdges, reflection, text, textRuns, textColor, textFontSize, textAlign, textVAlign, textBold, textItalic, textUnderline, textFontFamily, textMargin, textOutlineColor, textOutlineWidth, includeTextInSvg });
  return {
    name: preset.label || "図形",
    src: svgDataUrl(svg),
    mimeType: "image/svg+xml",
    extension: "svg",
    width,
    height,
    shape: {
      type: preset.type,
      label: preset.label || "図形",
      fillColor,
      ...(fillGradient ? { fillGradient } : {}),
      ...(fillPattern ? { fillPattern } : {}),
      outlineColor,
      outlineWidth,
      lineDash,
      lineCompound,
      lineCap,
      lineJoin,
      rotation,
      flipH,
      flipV,
      ...(shadow ? { shadow } : {}),
      ...(glow ? { glow } : {}),
      ...(softEdges ? { softEdges } : {}),
      ...(reflection ? { reflection } : {}),
      ...(points ? { points } : {}),
      ...(customPath ? { customPath } : {}),
      ...(isLine ? { lineStartArrow, lineEndArrow, ...(lineStartArrowType ? { lineStartArrowType } : {}), ...(lineEndArrowType ? { lineEndArrowType } : {}), ...(lineStartArrowSize ? { lineStartArrowSize } : {}), ...(lineEndArrowSize ? { lineEndArrowSize } : {}), ...(connections ? { connections } : {}) } : {}),
      ...(adjust ? { adjust } : {}),
      ...(text || textRuns.length || isTextBox ? { text, ...(textRuns.length ? { textRuns } : {}), textColor, textFontSize, textAlign, textVAlign, textBold, textItalic, textMargin, ...(textUnderline ? { textUnderline } : {}), ...(textFontFamily ? { textFontFamily } : {}), ...(textOutlineColor && textOutlineColor !== "none" ? { textOutlineColor, textOutlineWidth } : {}) } : {}),
    },
  };
}



function isLineShapeType(type) {
  return LINE_SHAPE_TYPES.has(type);
}



function isLineShapeObject(image) {
  return Boolean(image?.shape && isLineShapeType(image.shape.type));
}



function shapeCanEditPoints(image) {
  if (!image?.shape) return false;
  return isLineShapeObject(image) || shapeCustomPathHandles(image.shape.customPath).length > 0 || shapeAdjustmentHandles(image.shape.type, image.shape.adjust).length > 0;
}



function normalizeShapeOutlineWidth(value) {
  const width = Number(value);
  if (!Number.isFinite(width) || width <= 0) return DEFAULT_SHAPE_OUTLINE_WIDTH;
  return Math.max(0.6, Math.min(6, Math.round(width * 100) / 100));
}



function normalizeShapeTextOutlineWidth(value) {
  const width = Number(value);
  if (!Number.isFinite(width) || width <= 0) return 0.75;
  return Math.max(0.2, Math.min(3, Math.round(width * 100) / 100));
}



function normalizeShapeLineDash(value) {
  const raw = String(value || "solid");
  return SHAPE_LINE_DASH_PRESETS.some((preset) => preset.value === raw) ? raw : "solid";
}



function normalizeShapeLineCompound(value) {
  const raw = String(value || "single").trim();
  const map = {
    sng: "single",
    single: "single",
    dbl: "double",
    double: "double",
    thickThin: "thickThin",
    thinThick: "thinThick",
    tri: "triple",
    triple: "triple",
  };
  return SHAPE_LINE_COMPOUND_PRESETS.some((preset) => preset.value === (map[raw] || raw)) ? map[raw] || raw : "single";
}



function normalizeShapeLineCap(value) {
  const raw = String(value || "round").trim();
  const map = {
    flat: "butt",
    butt: "butt",
    sq: "square",
    square: "square",
    rnd: "round",
    round: "round",
  };
  return SHAPE_LINE_CAP_PRESETS.some((preset) => preset.value === (map[raw] || raw)) ? map[raw] || raw : "round";
}



function normalizeShapeLineJoin(value) {
  const raw = String(value || "round").trim();
  return SHAPE_LINE_JOIN_PRESETS.some((preset) => preset.value === raw) ? raw : "round";
}



function normalizeShapeArrowType(value) {
  const raw = String(value || "").trim().toLowerCase();
  const map = {
    triangle: "triangle",
    arrow: "triangle",
    stealth: "stealth",
    diamond: "diamond",
    oval: "oval",
    open: "open",
  };
  return map[raw] || (raw && raw !== "none" ? "triangle" : "");
}



function normalizeShapeFillPattern(value) {
  if (!value || typeof value !== "object") return null;
  const preset = SHAPE_FILL_PATTERN_PRESETS.some((item) => item.value === value.preset)
    ? value.preset
    : "pct20";
  if (!preset) return null;
  const foreground = cssColorForSvg(value.foreground) || "#4472c4";
  const background = cssColorForSvg(value.background) || "rgba(255,255,255,0)";
  return { preset, foreground, background };
}



function normalizeShapeGlow(value) {
  if (!value) return null;
  if (typeof value === "string") {
    const preset = SHAPE_GLOW_PRESETS.find((item) => item.value === value);
    return preset?.value ? { preset: preset.value, radius: preset.radius, color: preset.color } : null;
  }
  if (typeof value !== "object") return null;
  const radius = clampNumber(value.radius, 1, 24, 5);
  const color = cssColorForSvg(value.color) || "#5b9bd5";
  const preset = SHAPE_GLOW_PRESETS.find((item) => item.value && Math.abs(item.radius - radius) < 0.1 && item.color.toLowerCase() === color.toLowerCase())?.value || "custom";
  return { preset, radius, color };
}



function normalizeShapeSoftEdges(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return 0;
  return Math.max(0, Math.min(24, Math.round(number * 10) / 10));
}



function normalizeShapeReflection(value) {
  if (!value) return null;
  if (typeof value === "string") {
    if (!SHAPE_REFLECTION_PRESETS.some((item) => item.value === value) || !value) return null;
    return { preset: value };
  }
  if (typeof value !== "object") return null;
  const preset = SHAPE_REFLECTION_PRESETS.some((item) => item.value === value.preset) && value.preset ? value.preset : "tight";
  return {
    preset,
    blur: Math.max(0, Math.min(16, Math.round((Number(value.blur) || 0) * 10) / 10)),
    distance: Math.max(0, Math.min(48, Math.round((Number(value.distance) || 0) * 10) / 10)),
    alpha: Math.max(0.15, Math.min(0.85, Number(value.alpha) || 0.45)),
  };
}



function normalizeShapeCustomPath(value) {
  const commands = Array.isArray(value?.commands) ? value.commands : [];
  const normalized = commands
    .map((command) => {
      const cmd = String(command?.cmd || "").toUpperCase();
      if (cmd === "Z") return { cmd: "Z", points: [] };
      const expectedPoints = cmd === "C" ? 3 : cmd === "Q" ? 2 : ["M", "L"].includes(cmd) ? 1 : 0;
      if (!expectedPoints) return null;
      const points = (command.points || [])
        .slice(0, expectedPoints)
        .map((point) => ({
          x: clamp01(Number(point?.x)),
          y: clamp01(Number(point?.y)),
        }))
        .filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
      return points.length === expectedPoints ? { cmd, points } : null;
    })
    .filter(Boolean);
  if (!normalized.some((command) => command.cmd === "M")) return null;
  if (!normalized.some((command) => command.cmd !== "M" && command.cmd !== "Z")) return null;
  return { commands: normalized.slice(0, 300) };
}



function normalizeShapeArrowSize(value) {
  const raw = String(value || "").trim().toLowerCase();
  const map = {
    sm: "small",
    small: "small",
    med: "medium",
    medium: "medium",
    mid: "medium",
    lg: "large",
    large: "large",
  };
  return map[raw] || "medium";
}



function normalizeRotationDegrees(value) {
  const raw = Number(value);
  if (!Number.isFinite(raw)) return 0;
  const normalized = ((raw % 360) + 360) % 360;
  return Math.round(normalized * 10) / 10;
}



function normalizeShapeLinePoints(points) {
  const source = points && typeof points === "object" ? points : {};
  const x1 = clamp01(Number(source.x1));
  const y1 = clamp01(Number(source.y1));
  const x2 = clamp01(Number(source.x2));
  const y2 = clamp01(Number(source.y2));
  const normalized = {
    x1: Number.isFinite(x1) ? x1 : 0.04,
    y1: Number.isFinite(y1) ? y1 : 0.5,
    x2: Number.isFinite(x2) ? x2 : 0.96,
    y2: Number.isFinite(y2) ? y2 : 0.5,
  };
  if (Math.abs(normalized.x1 - normalized.x2) < 0.001 && Math.abs(normalized.y1 - normalized.y2) < 0.001) {
    normalized.x2 = Math.min(1, normalized.x1 + 0.2);
  }
  return normalized;
}



function normalizeLineShapeConnections(connections) {
  if (!connections || typeof connections !== "object") return null;
  const next = {};
  ["start", "end"].forEach((side) => {
    const item = connections[side];
    if (item?.imageId && item?.site) next[side] = { imageId: String(item.imageId), site: String(item.site) };
  });
  return next.start || next.end ? next : null;
}



function normalizeShapeGradient(gradient) {
  if (!gradient || typeof gradient !== "object") return null;
  const stops = (gradient.stops || [])
    .map((stop) => ({
      pos: clamp01(Number(stop?.pos)),
      color: cssColorForSvg(stop?.color),
    }))
    .filter((stop) => Number.isFinite(stop.pos) && stop.color)
    .sort((a, b) => a.pos - b.pos);
  if (stops.length < 2) return null;
  return {
    type: "linear",
    angle: normalizeRotationDegrees(gradient.angle),
    stops: stops.slice(0, 10),
  };
}



function normalizeShapeShadow(shadow) {
  if (!shadow) return null;
  if (typeof shadow === "string") return shadow;
  if (typeof shadow !== "object") return null;
  return {
    type: "outer",
    dx: Math.round((Number(shadow.dx) || 0) * 10) / 10,
    dy: Math.round((Number(shadow.dy) || 0) * 10) / 10,
    blur: Math.max(0, Math.round((Number(shadow.blur) || 0) * 10) / 10),
    color: cssColorForSvg(shadow.color) || "rgba(0,0,0,0.35)",
  };
}



function normalizeShapeTextMargin(margin, compact = false) {
  const source = margin && typeof margin === "object" ? margin : {};
  const fallback = compact
    ? { left: 5, right: 5, top: 4, bottom: 4 }
    : { left: 10, right: 10, top: 8, bottom: 8 };
  const next = {};
  ["left", "right", "top", "bottom"].forEach((side) => {
    next[side] = clampNumber(source[side], 0, 80, fallback[side]);
  });
  return next;
}



function normalizeShapeTextRuns(runs) {
  if (!Array.isArray(runs)) return [];
  const normalized = [];
  runs.slice(0, 500).forEach((run) => {
    const text = String(run?.text ?? "");
    if (!text) return;
    const next = { text };
    const color = cssColorForSvg(run.color);
    const fontSize = clampNumber(run.fontSize, 6, 72, Number.NaN);
    const fontFamily = String(run.fontFamily || "").trim();
    const underline = underlinePresetForValue(run.underline)?.value || "";
    if (color) next.color = color;
    if (Number.isFinite(fontSize)) next.fontSize = Math.round(fontSize * 10) / 10;
    if (fontFamily) next.fontFamily = fontFamily;
    if (typeof run.bold === "boolean") next.bold = run.bold;
    if (typeof run.italic === "boolean") next.italic = run.italic;
    if (underline) next.underline = underline;
    const previous = normalized[normalized.length - 1];
    if (previous && shapeTextRunStyleKey(previous) === shapeTextRunStyleKey(next)) {
      previous.text += next.text;
    } else {
      normalized.push(next);
    }
  });
  return normalized.some((run) => run.text.replace(/\n/g, "").length) ? normalized : [];
}



function shapeTextRunStyleKey(run) {
  return JSON.stringify({
    color: run.color || "",
    fontSize: run.fontSize || "",
    fontFamily: run.fontFamily || "",
    bold: typeof run.bold === "boolean" ? run.bold : "",
    italic: typeof run.italic === "boolean" ? run.italic : "",
    underline: run.underline || "",
  });
}



function shapePlainTextFromRuns(runs) {
  return normalizeShapeTextRuns(runs).map((run) => run.text).join("");
}



function clamp01(value) {
  return Number.isFinite(value) ? Math.max(0, Math.min(1, value)) : Number.NaN;
}



function clampNumber(value, min, max, fallback) {
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.max(min, Math.min(max, number));
}



function normalizeShapeAdjustments(type, adjust = {}) {
  const source = adjust && typeof adjust === "object" ? adjust : {};
  switch (type) {
    case "elbowConnector":
      return { bend: clampNumber(source.bend, 0, 1, 0.5), ...(["x", "y"].includes(source.axis) ? { axis: source.axis } : {}) };
    case "curvedConnector":
      return { curve: clampNumber(source.curve, 0.04, 0.96, 0.16) };
    case "roundRect":
      return { corner: clampNumber(source.corner, 0.02, 0.3, 0.1) };
    case "roundSnipRect":
      return {
        corner: clampNumber(source.corner, 0.02, 0.26, 0.14),
        cut: clampNumber(source.cut, 0.04, 0.42, 0.16),
      };
    case "snipRect":
      return { cut: clampNumber(source.cut, 0.04, 0.42, 0.16) };
    case "parallelogram":
      return { skew: clampNumber(source.skew, 0.04, 0.46, 0.22) };
    case "trapezoid":
      return { inset: clampNumber(source.inset, 0.04, 0.46, 0.24) };
    case "cross":
      return { thickness: clampNumber(source.thickness, 0.16, 0.76, 0.34) };
    case "can":
      return { depth: clampNumber(source.depth, 0.08, 0.42, 0.2) };
    case "rightBlockArrow":
    case "leftBlockArrow":
      return {
        thickness: clampNumber(source.thickness, 0.16, 0.82, 0.5),
        head: clampNumber(source.head, 0.22, 0.82, type === "leftBlockArrow" ? 0.4 : 0.6),
      };
    case "upBlockArrow":
    case "downBlockArrow":
      return {
        thickness: clampNumber(source.thickness, 0.16, 0.82, 0.5),
        head: clampNumber(source.head, 0.22, 0.82, type === "downBlockArrow" ? 0.48 : 0.52),
      };
    case "leftRightBlockArrow":
    case "quadArrow":
    case "leftRightUpArrow":
      return { thickness: clampNumber(source.thickness, 0.16, 0.72, 0.36) };
    case "notchedRightArrow":
      return {
        thickness: clampNumber(source.thickness, 0.16, 0.82, 0.42),
        head: clampNumber(source.head, 0.22, 0.82, 0.6),
      };
    case "bentUpArrow":
    case "uturnArrow":
      return { thickness: clampNumber(source.thickness, 0.16, 0.72, 0.38) };
    case "curvedRightArrow":
      return { thickness: clampNumber(source.thickness, 0.16, 0.72, 0.38) };
    case "flowData":
    case "flowManualInput":
      return { skew: clampNumber(source.skew, 0.04, 0.44, type === "flowData" ? 0.2 : 0.12) };
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
      return { inner: clampNumber(source.inner, 0.02, 1, defaultStarInnerRatio(type)) };
    case "calloutRect":
    case "calloutRoundRect":
    case "calloutOval":
    case "calloutCloud":
      return {
        tailX: clampNumber(source.tailX, -2, 3, defaultCalloutTailX(type)),
        tailY: clampNumber(source.tailY, -2, 3, defaultCalloutTailY(type)),
      };
    default:
      return null;
  }
}



function defaultStarInnerRatio(type) {
  const defaults = {
    star4: 0.62,
    star5: 0.42,
    star6: 0.39,
    star7: 0.4,
    star8: 0.45,
    star10: 0.48,
    star12: 0.5,
    star16: 0.54,
    star24: 0.6,
    star32: 0.64,
  };
  return defaults[type] || 0.45;
}



function starPointCountForType(type) {
  const match = String(type || "").match(/^star(\d+)$/);
  return match ? Number(match[1]) : 5;
}



function defaultCalloutTailX(type) {
  if (type === "calloutOval") return 0.34;
  if (type === "calloutCloud") return 0.37;
  return drawingCalloutPositionFromGuide(-20833);
}



function defaultCalloutTailY() {
  return drawingCalloutPositionFromGuide(62500);
}



function shapePresetForType(type) {
  if (type === CUSTOM_SHAPE_TYPE) return { type: CUSTOM_SHAPE_TYPE, label: "自由図形" };
  return RIBBON_SHAPE_PRESET_MAP.get(type) || RIBBON_SHAPE_PRESET_MAP.get(DEFAULT_SHAPE_TYPE) || null;
}



function shapeGroupForType(type) {
  return RIBBON_SHAPE_GROUPS.find((group) => group.shapes.some((shape) => shape.type === type)) || null;
}



function setCurrentRibbonShape(type) {
  const preset = shapePresetForType(type);
  if (!preset) return;
  state.currentShapeType = preset.type;
  const recent = [preset.type, ...(state.recentShapeTypes || [])].filter(Boolean);
  state.recentShapeTypes = [...new Set(recent)].filter((item) => shapePresetForType(item)).slice(0, 8);
}



function roundRectCornerScale(width, height) {
  const w = Math.max(1, Number(width) || 1);
  const h = Math.max(1, Number(height) || 1);
  const side = Math.min(w, h);
  return { x: side / w, y: side / h };
}



function shapeSvgMarkup(type, options = {}) {
  const width = Number(options.width) || 150;
  const height = Number(options.height) || 100;
  const fillGradient = normalizeShapeGradient(options.fillGradient);
  const fillPattern = normalizeShapeFillPattern(options.fillPattern);
  const fill = fillPattern ? "url(#shapeFillPattern)" : fillGradient ? "url(#shapeFillGradient)" : options.fillColor === "none" ? "none" : cssColorForSvg(options.fillColor) || "#eaf3ff";
  const stroke = options.outlineColor === "none" ? "none" : cssColorForSvg(options.outlineColor) || "#4472c4";
  const strokeWidth = normalizeShapeOutlineWidth(options.outlineWidth);
  const lineDash = normalizeShapeLineDash(options.lineDash);
  const lineCap = normalizeShapeLineCap(options.lineCap);
  const lineJoin = normalizeShapeLineJoin(options.lineJoin);
  const dashArray = shapeLineDashArray(lineDash, strokeWidth);
  const dashAttr = dashArray ? ` stroke-dasharray="${dashArray}"` : "";
  const dark = "#1f2937";
  const adjust = normalizeShapeAdjustments(type, options.adjust);
  const customPath = normalizeShapeCustomPath(options.customPath);
  const shapeAttrs = `fill="${escapeXml(fill)}" stroke="${escapeXml(stroke)}" stroke-width="${stroke === "none" ? 0 : strokeWidth}" stroke-linecap="${escapeXml(lineCap)}" stroke-linejoin="${escapeXml(lineJoin)}"${dashAttr} vector-effect="non-scaling-stroke"`;
  const lineAttrs = `fill="none" stroke="${escapeXml(stroke)}" stroke-width="${stroke === "none" ? 0 : strokeWidth}" stroke-linecap="${escapeXml(lineCap)}" stroke-linejoin="${escapeXml(lineJoin)}"${dashAttr} vector-effect="non-scaling-stroke"`;
  const points = normalizeShapeLinePoints(options.points);
  const p1 = svgPointFromLinePoint(points.x1, points.y1);
  const p2 = svgPointFromLinePoint(points.x2, points.y2);
  const lineStartArrow = Boolean(options.lineStartArrow ?? type === "doubleArrow");
  const lineEndArrow = Boolean(options.lineEndArrow ?? ["arrow", "doubleArrow", "elbowConnector", "curvedConnector"].includes(type));
  const lineStartArrowType = normalizeShapeArrowType(options.lineStartArrowType || "triangle");
  const lineEndArrowType = normalizeShapeArrowType(options.lineEndArrowType || "triangle");
  const lineStartArrowSize = normalizeShapeArrowSize(options.lineStartArrowSize);
  const lineEndArrowSize = normalizeShapeArrowSize(options.lineEndArrowSize);
  if (isLineShapeType(type)) {
    return shapeLineSvgMarkup(type, {
      width,
      height,
      stroke,
      strokeWidth,
      lineDash,
      lineCap,
      lineJoin,
      dashAttr,
      adjust,
      points,
      lineStartArrow,
      lineEndArrow,
      lineStartArrowType,
      lineEndArrowType,
      lineStartArrowSize,
      lineEndArrowSize,
      lineCompound: options.lineCompound,
      shadow: options.shadow,
      glow: options.glow,
      softEdges: options.softEdges,
      reflection: options.reflection,
    });
  }
  let body = "";
  switch (type) {
    case "line":
      body = shapeLineSvgWithArrows([p1, p2], {
        lineAttrs,
        stroke,
        strokeWidth,
        width,
        height,
        lineStartArrow,
        lineEndArrow,
        lineStartArrowType,
        lineEndArrowType,
        lineStartArrowSize,
        lineEndArrowSize,
      });
      break;
    case "arrow":
      body = shapeLineSvgWithArrows([p1, p2], {
        lineAttrs,
        stroke,
        strokeWidth,
        width,
        height,
        lineStartArrow,
        lineEndArrow,
        lineStartArrowType,
        lineEndArrowType,
        lineStartArrowSize,
        lineEndArrowSize,
      });
      break;
    case "doubleArrow":
      body = shapeLineSvgWithArrows([p1, p2], {
        lineAttrs,
        stroke,
        strokeWidth,
        width,
        height,
        lineStartArrow,
        lineEndArrow,
        lineStartArrowType,
        lineEndArrowType,
        lineStartArrowSize,
        lineEndArrowSize,
      });
      break;
    case "elbowConnector":
      {
        const route = elbowConnectorRoutePoints(p1, p2, adjust);
        body = shapePolylineSvgWithArrows(route, {
          lineAttrs,
          stroke,
          strokeWidth,
          width,
          height,
          lineStartArrow,
          lineEndArrow,
          lineStartArrowType,
          lineEndArrowType,
          lineStartArrowSize,
          lineEndArrowSize,
        });
      }
      break;
    case "curvedConnector":
      {
        const curveOffset = roundSvgNumber((0.5 - (adjust?.curve ?? 0.16)) * 52);
        const c1 = { x: roundSvgNumber(p1.x + (p2.x - p1.x) * 0.38), y: roundSvgNumber(p1.y + curveOffset) };
        const c2 = { x: roundSvgNumber(p1.x + (p2.x - p1.x) * 0.62), y: roundSvgNumber(p2.y + curveOffset) };
        body = shapeBezierSvgWithArrows(p1, c1, c2, p2, {
          lineAttrs,
          stroke,
          strokeWidth,
          width,
          height,
          lineStartArrow,
          lineEndArrow,
          lineStartArrowType,
          lineEndArrowType,
          lineStartArrowSize,
          lineEndArrowSize,
        });
      }
      break;
    case "freeform":
      body = shapeFreeformSvgWithArrows(p1, p2, {
        lineAttrs,
        stroke,
        strokeWidth,
        width,
        height,
        lineStartArrow,
        lineEndArrow,
        lineStartArrowType,
        lineEndArrowType,
        lineStartArrowSize,
        lineEndArrowSize,
      });
      break;
    case "customPath":
      body = customPath
        ? `<path d="${escapeAttr(shapeCustomPathSvgD(customPath))}" ${shapeAttrs}/>`
        : `<rect x="0.8" y="0.8" width="98.4" height="58.4" ${shapeAttrs}/>`;
      break;
    case "roundRect":
      {
        const cornerScale = roundRectCornerScale(width, height);
        body = `<rect x="0.2" y="0.2" width="99.6" height="59.6" rx="${roundSvgNumber(adjust.corner * 100 * cornerScale.x)}" ry="${roundSvgNumber(adjust.corner * 60 * cornerScale.y)}" ${shapeAttrs}/>`;
      }
      break;
    case "snipRect":
      {
        const cut = roundSvgNumber(adjust.cut * 100);
        body = `<path d="M0.2 0.2 H${roundSvgNumber(100 - cut)} L99.8 ${roundSvgNumber(cut * 0.6)} V59.8 H${cut} L0.2 ${roundSvgNumber(60 - cut * 0.6)} Z" ${shapeAttrs}/>`;
      }
      break;
    case "roundSnipRect":
      {
        const corner = roundSvgNumber(adjust.corner * 100);
        const cut = roundSvgNumber(adjust.cut * 100);
        body = `<path d="M${corner} 0.2 H${roundSvgNumber(100 - cut)} L99.8 ${roundSvgNumber(cut * 0.6)} V${roundSvgNumber(60 - corner)} Q99.8 59.8 ${roundSvgNumber(100 - corner)} 59.8 H${corner} Q0.2 59.8 0.2 ${roundSvgNumber(60 - corner)} V${corner} Q0.2 0.2 ${corner} 0.2 Z" ${shapeAttrs}/>`;
      }
      break;
    case "frame":
      body = `<path d="M0.2 0.2 H99.8 V59.8 H0.2 Z M20 18 V42 H80 V18 Z" fill="${escapeXml(fill)}" stroke="${escapeXml(stroke)}" stroke-width="${stroke === "none" ? 0 : strokeWidth}" fill-rule="evenodd" vector-effect="non-scaling-stroke"/>`;
      break;
    case "textBox":
      body = `<rect x="0.2" y="0.2" width="99.6" height="59.6" ${shapeAttrs}/>`;
      break;
    case "oval":
      body = `<ellipse cx="50" cy="30" rx="49.8" ry="29.8" ${shapeAttrs}/>`;
      break;
    case "triangle":
      body = `<polygon points="50,0.2 99.8,59.8 0.2,59.8" ${shapeAttrs}/>`;
      break;
    case "rightTriangle":
      body = `<polygon points="0.2,0.2 99.8,59.8 0.2,59.8" ${shapeAttrs}/>`;
      break;
    case "diamond":
      body = `<polygon points="50,0.2 99.8,30 50,59.8 0.2,30" ${shapeAttrs}/>`;
      break;
    case "pentagon":
      body = `<polygon points="${regularPolygonPoints(50, 31, 44, 5, -90, { fit: true })}" ${shapeAttrs}/>`;
      break;
    case "hexagon":
      body = `<polygon points="${regularPolygonPoints(50, 30, 43, 6, 30, { fit: true })}" ${shapeAttrs}/>`;
      break;
    case "heptagon":
      body = `<polygon points="${regularPolygonPoints(50, 30, 43, 7, -90, { fit: true })}" ${shapeAttrs}/>`;
      break;
    case "octagon":
      body = `<polygon points="29,0.2 71,0.2 99.8,18 99.8,42 71,59.8 29,59.8 0.2,42 0.2,18" ${shapeAttrs}/>`;
      break;
    case "decagon":
      body = `<polygon points="${regularPolygonPoints(50, 30, 43, 10, -90, { fit: true })}" ${shapeAttrs}/>`;
      break;
    case "dodecagon":
      body = `<polygon points="${regularPolygonPoints(50, 30, 43, 12, -90, { fit: true })}" ${shapeAttrs}/>`;
      break;
    case "parallelogram":
      {
        const skew = roundSvgNumber(adjust.skew * 100);
        body = `<polygon points="${skew},0.2 99.8,0.2 ${roundSvgNumber(100 - skew)},59.8 0.2,59.8" ${shapeAttrs}/>`;
      }
      break;
    case "trapezoid":
      {
        const inset = roundSvgNumber(adjust.inset * 100);
        body = `<polygon points="${inset},0.2 ${roundSvgNumber(100 - inset)},0.2 99.8,59.8 0.2,59.8" ${shapeAttrs}/>`;
      }
      break;
    case "cross":
      {
        const x1 = roundSvgNumber(((1 - adjust.thickness) / 2) * 100);
        const x2 = roundSvgNumber(100 - x1);
        const y1 = roundSvgNumber(((1 - adjust.thickness) / 2) * 60);
        const y2 = roundSvgNumber(60 - y1);
        body = `<path d="M${x1} 0.2 H${x2} V${y1} H99.8 V${y2} H${x2} V59.8 H${x1} V${y2} H0.2 V${y1} H${x1} Z" ${shapeAttrs}/>`;
      }
      break;
    case "can":
      {
        const depth = roundSvgNumber(adjust.depth * 60);
        const bottom = roundSvgNumber(60 - depth);
        const ellipseRy = roundSvgNumber(Math.max(4, depth - 0.2));
        body = `<path d="M0.2 ${depth} C0.2 0.2 99.8 0.2 99.8 ${depth} V${bottom} C99.8 ${roundSvgNumber(60 + depth)} 0.2 ${roundSvgNumber(60 + depth)} 0.2 ${bottom} Z" ${shapeAttrs}/><ellipse cx="50" cy="${depth}" rx="49.8" ry="${ellipseRy}" fill="#f8fbff" stroke="${escapeXml(stroke)}" stroke-width="${stroke === "none" ? 0 : strokeWidth}" vector-effect="non-scaling-stroke"/><path d="M0.2 ${bottom} C0.2 ${roundSvgNumber(60 + depth)} 99.8 ${roundSvgNumber(60 + depth)} 99.8 ${bottom}" ${lineAttrs}/>`;
      }
      break;
    case "cube":
      body = `<path d="M0.2 13 L22 0.2 H99.8 V46 L78 59.8 H0.2 Z" ${shapeAttrs}/><path d="M0.2 13 H78 L99.8 0.2 M78 13 V59.8" ${lineAttrs}/>`;
      break;
    case "bevel":
      body = `<path d="M0.2 0.2 H99.8 V59.8 H0.2 Z M14 10 H86 V50 H14 Z" fill="${escapeXml(fill)}" stroke="${escapeXml(stroke)}" stroke-width="${stroke === "none" ? 0 : strokeWidth}" fill-rule="evenodd" vector-effect="non-scaling-stroke"/><path d="M0.2 0.2 L14 10 M99.8 0.2 L86 10 M99.8 59.8 L86 50 M0.2 59.8 L14 50" ${lineAttrs}/>`;
      break;
    case "donut":
      body = `<path d="M50 0.2 C77.5 0.2 99.8 13.6 99.8 30 S77.5 59.8 50 59.8 0.2 46.4 0.2 30 22.5 0.2 50 0.2 Z M50 16.8 C65.5 16.8 78.2 22.7 78.2 30 S65.5 43.2 50 43.2 21.8 37.3 21.8 30 34.5 16.8 50 16.8 Z" fill="${escapeXml(fill)}" stroke="${escapeXml(stroke)}" stroke-width="${stroke === "none" ? 0 : strokeWidth}" fill-rule="evenodd" vector-effect="non-scaling-stroke"/>`;
      break;
    case "noSymbol":
      body = `<ellipse cx="50" cy="30" rx="49.8" ry="29.8" ${shapeAttrs}/><path d="M17 50 L83 10" ${lineAttrs}/>`;
      break;
    case "blockArc":
      body = `<path d="M12 58 C2 34 16 7 43 1 C70 -5 96 10 99 35 L78 37 C76 25 62 16 47 20 C32 23 24 38 30 51 Z" ${shapeAttrs}/>`;
      break;
    case "foldedCorner":
      body = `<path d="M0.2 0.2 H86 L99.8 14 V59.8 H0.2 Z" ${shapeAttrs}/><path d="M86 0.2 V14 H99.8" ${lineAttrs}/>`;
      break;
    case "smileyFace":
      body = `<ellipse cx="50" cy="30" rx="49.8" ry="29.8" ${shapeAttrs}/><circle cx="33" cy="23" r="3.2" fill="${escapeXml(stroke === "none" ? dark : stroke)}"/><circle cx="67" cy="23" r="3.2" fill="${escapeXml(stroke === "none" ? dark : stroke)}"/><path d="M30 38 C39 49 61 49 70 38" ${lineAttrs}/>`;
      break;
    case "heart":
      body = `<path d="M50 56 C32 43 5 29 8 13 C11 -4 36 -2 50 15 C64 -2 89 -4 92 13 C95 29 68 43 50 56 Z" ${shapeAttrs}/>`;
      break;
    case "lightningBolt":
      body = `<polygon points="55,0.2 18,34 43,34 34,59.8 82,22 55,22" ${shapeAttrs}/>`;
      break;
    case "sun":
      body = `<polygon points="${starPoints(50, 30, 31, 22, 16, { fit: true })}" ${shapeAttrs}/><ellipse cx="50" cy="30" rx="24" ry="14.4" ${shapeAttrs}/>`;
      break;
    case "moon":
      body = `<path d="M70 3 C49 8 35 22 35 37 C35 48 45 56 59 58 C32 59 8 46 8 30 C8 12 35 -1 70 3 Z" ${shapeAttrs}/>`;
      break;
    case "cloud":
      body = `<path d="M0.2 40 C0.2 31 7 25 16 25 C17 13 28 6 40 12 C47 1 65 2 72 15 C88 14 99.8 24 99.8 38 C99.8 50 88 59.8 73 59.8 H19 C8 59.8 0.2 51 0.2 40 Z" ${shapeAttrs}/>`;
      break;
    case "arc":
      body = `<path d="M18 51 C20 16 44 2 82 9" ${lineAttrs}/>`;
      break;
    case "pie":
      body = `<path d="M50 30 L50 0.2 A49.8 29.8 0 1 1 8 46 Z" ${shapeAttrs}/>`;
      break;
    case "chord":
      body = `<path d="M0.2 30 A49.8 29.8 0 0 1 99.8 30 L80 59.8 H20 Z" ${shapeAttrs}/>`;
      break;
    case "teardrop":
      body = `<path d="M50 0.2 C83 0.2 99.8 18 99.8 36 C99.8 51 84 59.8 65 59.8 C47 59.8 35 49 35 35 C35 22 42 11 50 0.2 Z" ${shapeAttrs}/>`;
      break;
    case "plaque":
      body = `<path d="M18 0.2 C18 10 10 18 0.2 18 V42 C10 42 18 50 18 59.8 H82 C82 50 90 42 99.8 42 V18 C90 18 82 10 82 0.2 Z" ${shapeAttrs}/>`;
      break;
    case "corner":
      body = `<path d="M0.2 0.2 H99.8 V18 H18 V59.8 H0.2 Z" ${shapeAttrs}/>`;
      break;
    case "halfFrame":
      body = `<path d="M0.2 0.2 H99.8 V18 H18 V59.8 H0.2 Z" ${shapeAttrs}/><path d="M18 18 H99.8 M18 18 V59.8" ${lineAttrs}/>`;
      break;
    case "diagStripe":
      body = `<path d="M0.2 59.8 L60 0.2 H99.8 L40 59.8 Z" ${shapeAttrs}/>`;
      break;
    case "leftBracket":
      body = `<path d="M82 0.2 H28 V59.8 H82" ${lineAttrs}/>`;
      break;
    case "rightBracket":
      body = `<path d="M18 0.2 H72 V59.8 H18" ${lineAttrs}/>`;
      break;
    case "bracketPair":
      body = `<path d="M38 0.2 H18 V59.8 H38 M62 0.2 H82 V59.8 H62" ${lineAttrs}/>`;
      break;
    case "leftBrace":
      body = `<path d="M75 0.2 C48 0.2 52 17 45 21 C40 24 31 23 27 30 C31 37 40 36 45 39 C52 43 48 59.8 75 59.8" ${lineAttrs}/>`;
      break;
    case "rightBrace":
      body = `<path d="M25 0.2 C52 0.2 48 17 55 21 C60 24 69 23 73 30 C69 37 60 36 55 39 C48 43 52 59.8 25 59.8" ${lineAttrs}/>`;
      break;
    case "bracePair":
      body = `<path d="M43 0.2 C23 0.2 26 17 21 21 C17 24 10 23 8 30 C10 37 17 36 21 39 C26 43 23 59.8 43 59.8 M57 0.2 C77 0.2 74 17 79 21 C83 24 90 23 92 30 C90 37 83 36 79 39 C74 43 77 59.8 57 59.8" ${lineAttrs}/>`;
      break;
    case "rightBlockArrow":
      {
        const y1 = roundSvgNumber(30 - adjust.thickness * 30);
        const y2 = roundSvgNumber(30 + adjust.thickness * 30);
        const head = roundSvgNumber(adjust.head * 100);
        body = `<path d="M0.2 ${y1} H${head} V0.2 L99.8 30 L${head} 59.8 V${y2} H0.2 Z" ${shapeAttrs}/>`;
      }
      break;
    case "leftBlockArrow":
      {
        const y1 = roundSvgNumber(30 - adjust.thickness * 30);
        const y2 = roundSvgNumber(30 + adjust.thickness * 30);
        const head = roundSvgNumber(adjust.head * 100);
        body = `<path d="M99.8 ${y1} H${head} V0.2 L0.2 30 L${head} 59.8 V${y2} H99.8 Z" ${shapeAttrs}/>`;
      }
      break;
    case "upBlockArrow":
      {
        const x1 = roundSvgNumber(50 - adjust.thickness * 50);
        const x2 = roundSvgNumber(50 + adjust.thickness * 50);
        const head = roundSvgNumber(adjust.head * 60);
        body = `<path d="M${x1} 59.8 V${head} H0.2 L50 0.2 L99.8 ${head} H${x2} V59.8 Z" ${shapeAttrs}/>`;
      }
      break;
    case "downBlockArrow":
      {
        const x1 = roundSvgNumber(50 - adjust.thickness * 50);
        const x2 = roundSvgNumber(50 + adjust.thickness * 50);
        const head = roundSvgNumber(adjust.head * 60);
        body = `<path d="M${x1} 0.2 V${head} H0.2 L50 59.8 L99.8 ${head} H${x2} V0.2 Z" ${shapeAttrs}/>`;
      }
      break;
    case "leftRightBlockArrow":
      {
        const y1 = roundSvgNumber(30 - adjust.thickness * 30);
        const y2 = roundSvgNumber(30 + adjust.thickness * 30);
        body = `<path d="M0.2 30 L28 0.2 V${y1} H72 V0.2 L99.8 30 L72 59.8 V${y2} H28 V59.8 Z" ${shapeAttrs}/>`;
      }
      break;
    case "leftRightUpArrow":
      {
        const y1 = roundSvgNumber(30 - adjust.thickness * 30);
        const y2 = roundSvgNumber(30 + adjust.thickness * 30);
        const x1 = roundSvgNumber(50 - adjust.thickness * 50);
        const x2 = roundSvgNumber(50 + adjust.thickness * 50);
        body = `<path d="M50 0.2 L${x2} ${y1} H72 V0.2 L99.8 30 L72 59.8 V${y2} H28 V59.8 L0.2 30 L28 0.2 V${y1} H${x1} Z" ${shapeAttrs}/>`;
      }
      break;
    case "quadArrow":
      {
        const halfY = roundSvgNumber(adjust.thickness * 30);
        const halfX = roundSvgNumber(adjust.thickness * 50);
        const y1 = roundSvgNumber(30 - halfY);
        const y2 = roundSvgNumber(30 + halfY);
        const x1 = roundSvgNumber(50 - halfX);
        const x2 = roundSvgNumber(50 + halfX);
        body = `<path d="M50 0.2 L${roundSvgNumber(50 + halfX * 0.75)} ${y1} H${x2} V${roundSvgNumber(30 - halfY * 0.45)} H${roundSvgNumber(100 - halfX * 0.75)} V${roundSvgNumber(30 - halfY)} L99.8 30 L${roundSvgNumber(100 - halfX * 0.75)} ${roundSvgNumber(30 + halfY)} V${roundSvgNumber(30 + halfY * 0.45)} H${x2} V${y2} L50 59.8 L${x1} ${y2} V${roundSvgNumber(30 + halfY * 0.45)} H${roundSvgNumber(halfX * 0.75)} V${roundSvgNumber(30 + halfY)} L0.2 30 L${roundSvgNumber(halfX * 0.75)} ${roundSvgNumber(30 - halfY)} V${roundSvgNumber(30 - halfY * 0.45)} H${x1} V${y1} Z" ${shapeAttrs}/>`;
      }
      break;
    case "notchedRightArrow":
      {
        const y1 = roundSvgNumber(30 - adjust.thickness * 30);
        const y2 = roundSvgNumber(30 + adjust.thickness * 30);
        const head = roundSvgNumber(adjust.head * 100);
        body = `<path d="M0.2 ${y1} H${head} V0.2 L99.8 30 L${head} 59.8 V${y2} H0.2 L13 30 Z" ${shapeAttrs}/>`;
      }
      break;
    case "bentUpArrow":
      {
        const shaft = roundSvgNumber(adjust.thickness * 50);
        const left = roundSvgNumber(50 - shaft / 2);
        const right = roundSvgNumber(50 + shaft / 2);
        body = `<path d="M0.2 59.8 V${roundSvgNumber(60 - shaft)} H${left} V24 H25 L50 0.2 L75 24 H${right} V59.8 Z" ${shapeAttrs}/>`;
      }
      break;
    case "uturnArrow":
      {
        const thick = roundSvgNumber(adjust.thickness * 30);
        body = `<path d="M78 59.8 V27 C78 14 68 6 52 6 H34 V0.2 L0.2 20 L34 39.8 V31 H52 C57 31 60 34 60 39 V59.8 Z M${roundSvgNumber(78 - thick)} 59.8 V39 C${roundSvgNumber(78 - thick)} 24 66 20 52 20 H34 V31 H52 C67 31 78 40 78 55 V59.8 Z" fill="${escapeXml(fill)}" stroke="${escapeXml(stroke)}" stroke-width="${stroke === "none" ? 0 : strokeWidth}" fill-rule="evenodd" vector-effect="non-scaling-stroke"/>`;
      }
      break;
    case "curvedRightArrow":
      {
        const innerTop = roundSvgNumber(60 - adjust.thickness * 88);
        const outerBottom = roundSvgNumber(60 - adjust.thickness * 24);
        body = `<path d="M0.2 ${outerBottom} C9 15 43 0.2 70 16 V0.2 L99.8 30 L70 59.8 V44 C48 ${innerTop} 25 ${roundSvgNumber(innerTop + 4)} 8 59.8 Z" ${shapeAttrs}/>`;
      }
      break;
    case "leftUpArrow":
      body = `<path d="M0.2 0.2 H52 V14 H25 L99.8 54 L88 59.8 L14 20 V47 H0.2 Z" ${shapeAttrs}/>`;
      break;
    case "stripedRightArrow":
      body = `<path d="M22 18 H66 V0.2 L99.8 30 L66 59.8 V42 H22 Z" ${shapeAttrs}/><path d="M0.2 18 H8 M0.2 30 H13 M0.2 42 H8" ${lineAttrs}/>`;
      break;
    case "homePlate":
      body = `<polygon points="0.2,0.2 75,0.2 99.8,30 75,59.8 0.2,59.8" ${shapeAttrs}/>`;
      break;
    case "chevron":
      body = `<polygon points="0.2,0.2 68,0.2 99.8,30 68,59.8 0.2,59.8 31,30" ${shapeAttrs}/>`;
      break;
    case "leftRightCircularArrow":
      body = `<path d="M20 43 C3 30 7 10 31 6 C50 2 67 10 74 24 H86 L72 0.2 L55 24 H66 C60 14 47 11 34 14 C20 18 17 29 27 37 L39 29 V59.8 L0.2 54 Z" ${shapeAttrs}/><path d="M80 17 C97 30 93 50 69 54 C50 58 33 50 26 36 H14 L28 59.8 L45 36 H34 C40 46 53 49 66 46 C80 42 83 31 73 23 L61 31 V0.2 L99.8 6 Z" ${shapeAttrs}/>`;
      break;
    case "mathPlus":
    case "mathMinus":
    case "mathMultiply":
    case "mathDivide":
    case "mathEqual":
    case "mathNotEqual":
      body = `<rect x="0.2" y="0.2" width="99.6" height="59.6" rx="4" ${shapeAttrs}/><text x="50" y="40" text-anchor="middle" font-family="Aptos, Arial, sans-serif" font-size="34" font-weight="700" fill="${dark}">${mathShapeText(type)}</text>`;
      break;
    case "flowProcess":
      body = `<rect x="0.2" y="0.2" width="99.6" height="59.6" ${shapeAttrs}/>`;
      break;
    case "flowDecision":
      body = `<polygon points="50,0.2 99.8,30 50,59.8 0.2,30" ${shapeAttrs}/>`;
      break;
    case "flowData":
      {
        const skew = roundSvgNumber(adjust.skew * 100);
        body = `<polygon points="${skew},0.2 99.8,0.2 ${roundSvgNumber(100 - skew)},59.8 0.2,59.8" ${shapeAttrs}/>`;
      }
      break;
    case "flowDocument":
      body = `<path d="M0.2 0.2 H99.8 V48 C82 55 70 52 58 49 C50 47 48 59.8 36 59.8 C24 59.8 12 54 0.2 52 Z" ${shapeAttrs}/>`;
      break;
    case "flowMultiDocument":
      body = `<path d="M12 8 H99.8 V50 C83 56 72 53 61 50 C53 48 51 59.8 40 59.8 C28 59.8 18 55 12 52 Z" ${shapeAttrs}/><path d="M6 4 H94 V45 M0.2 0.2 H88 V40" ${lineAttrs}/>`;
      break;
    case "flowTerminator":
      body = `<rect x="0.2" y="0.2" width="99.6" height="59.6" rx="29.8" ${shapeAttrs}/>`;
      break;
    case "flowPredefinedProcess":
      body = `<rect x="0.2" y="0.2" width="99.6" height="59.6" ${shapeAttrs}/><path d="M16 0.2 V59.8 M84 0.2 V59.8" ${lineAttrs}/>`;
      break;
    case "flowStoredData":
      body = `<path d="M22 0.2 H82 C93 0.2 99.8 12 99.8 30 C99.8 48 93 59.8 82 59.8 H22 C9 59.8 0.2 48 0.2 30 C0.2 12 9 0.2 22 0.2 Z" ${shapeAttrs}/><path d="M22 0.2 C34 14 34 46 22 59.8" ${lineAttrs}/>`;
      break;
    case "flowInternalStorage":
      body = `<rect x="0.2" y="0.2" width="99.6" height="59.6" ${shapeAttrs}/><path d="M20 0.2 V59.8 M0.2 18 H99.8" ${lineAttrs}/>`;
      break;
    case "flowManualInput":
      body = `<path d="M0.2 ${roundSvgNumber(adjust.skew * 60)} L99.8 0.2 V59.8 H0.2 Z" ${shapeAttrs}/>`;
      break;
    case "flowManualOperation":
      body = `<polygon points="0.2,0.2 99.8,0.2 82,59.8 18,59.8" ${shapeAttrs}/>`;
      break;
    case "flowPreparation":
      body = `<polygon points="18,0.2 82,0.2 99.8,30 82,59.8 18,59.8 0.2,30" ${shapeAttrs}/>`;
      break;
    case "flowConnector":
      body = `<ellipse cx="50" cy="30" rx="29.8" ry="29.8" ${shapeAttrs}/>`;
      break;
    case "flowOffpageConnector":
      body = `<polygon points="0.2,0.2 99.8,0.2 99.8,41 50,59.8 0.2,41" ${shapeAttrs}/>`;
      break;
    case "flowDelay":
      body = `<path d="M0.2 0.2 H62 C84 0.2 99.8 13.5 99.8 30 C99.8 46.5 84 59.8 62 59.8 H0.2 Z" ${shapeAttrs}/>`;
      break;
    case "flowDisplay":
      body = `<path d="M0.2 30 L18 0.2 H78 C91 0.2 99.8 13.5 99.8 30 C99.8 46.5 91 59.8 78 59.8 H18 Z" ${shapeAttrs}/>`;
      break;
    case "flowOr":
      body = `<ellipse cx="50" cy="30" rx="29.8" ry="29.8" ${shapeAttrs}/><path d="M50 8 V52 M26 30 H74" ${lineAttrs}/>`;
      break;
    case "flowSummingJunction":
      body = `<ellipse cx="50" cy="30" rx="29.8" ry="29.8" ${shapeAttrs}/><path d="M30 12 L70 48 M70 12 L30 48" ${lineAttrs}/>`;
      break;
    case "star4":
      body = `<polygon points="${star4Points(adjust.inner)}" ${shapeAttrs}/>`;
      break;
    case "star5":
    case "star6":
    case "star7":
    case "star8":
    case "star10":
    case "star12":
    case "star16":
    case "star24":
    case "star32":
      body = `<polygon points="${starPoints(50, 30, 31, adjust.inner * 31, starPointCountForType(type), { fit: true })}" ${shapeAttrs}/>`;
      break;
    case "seal":
      body = `<polygon points="${explosionSealPoints()}" ${shapeAttrs}/>`;
      break;
    case "explosion":
      body = `<polygon points="${explosionBurstPoints()}" ${shapeAttrs}/>`;
      break;
    case "ribbon":
      body = `<path d="M0.2 0.2 H99.8 V45 H66 L50 59.8 L34 45 H0.2 Z" ${shapeAttrs}/><path d="M34 45 V18 H66 V45" ${lineAttrs}/>`;
      break;
    case "wave":
      body = `<path d="M0.2 36 C12 16 18 0.2 30 0.2 C43 0.2 42 59.8 55 59.8 C70 59.8 76 0.2 99.8 32 V59.8 H0.2 Z" ${shapeAttrs}/>`;
      break;
    case "calloutRoundRect":
      body = calloutShapePath(type, adjust, shapeAttrs);
      break;
    case "calloutOval":
      body = calloutShapePath(type, adjust, shapeAttrs);
      break;
    case "calloutCloud":
      body = calloutShapePath(type, adjust, shapeAttrs);
      break;
    case "calloutRect":
      body = calloutShapePath(type, adjust, shapeAttrs);
      break;
    case "rect":
    default:
      body = `<rect x="0" y="0" width="100" height="60" ${shapeAttrs}/>`;
      break;
  }
  body = shapeCompoundLineSvgBody(body, { lineCompound: options.lineCompound, stroke, strokeWidth, fill, isLine: isLineShapeType(type) });
  const effects = shapeEffectsSvg(options, width, height);
  const textBody = isLineShapeType(type) || options.includeTextInSvg === false ? "" : shapeTextSvgMarkup(options.text, {...options,type});
  const gradientDef = fillGradient ? shapeGradientDefSvg(fillGradient) : "";
  const patternDef = fillPattern ? shapePatternDefSvg(fillPattern) : "";
  const viewBox = isCalloutShapeType(type) ? calloutViewBoxAttr(adjust, type) : "0 0 100 60";
  const fittedBody = shapeBodyFittedToViewBox(body, viewBox, width, height, stroke, strokeWidth, options.lineCompound);
  const renderedBody = effects.filterAttr ? `<g ${effects.filterAttr}>${fittedBody}</g>` : fittedBody;
  const reflectionBody = effects.reflection ? shapeReflectionSvgMarkup(fittedBody, effects.reflection) : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="${viewBox}" preserveAspectRatio="none" overflow="visible">${gradientDef}${patternDef}${effects.defs}${reflectionBody}${renderedBody}${textBody}</svg>`;
}



function shapeLineSvgMarkup(type, options = {}) {
  const width = Math.max(1, Number(options.width) || 180);
  const height = Math.max(1, Number(options.height) || 70);
  const stroke = options.stroke === "none" ? "none" : cssColorForSvg(options.stroke) || "#4472c4";
  const strokeWidth = normalizeShapeOutlineWidth(options.strokeWidth);
  const lineCap = normalizeShapeLineCap(options.lineCap);
  const lineJoin = normalizeShapeLineJoin(options.lineJoin);
  const dashArray = shapeLineDashArray(options.lineDash, strokeWidth);
  const dashAttr = dashArray ? ` stroke-dasharray="${dashArray}"` : "";
  const lineAttrs = `fill="none" stroke="${escapeXml(stroke)}" stroke-width="${stroke === "none" ? 0 : strokeWidth}" stroke-linecap="${escapeXml(lineCap)}" stroke-linejoin="${escapeXml(lineJoin)}"${dashAttr} vector-effect="non-scaling-stroke"`;
  const points = normalizeShapeLinePoints(options.points);
  const p1 = linePointToPixelPoint(points.x1, points.y1, width, height);
  const p2 = linePointToPixelPoint(points.x2, points.y2, width, height);
  const arrowOptions = {
    lineAttrs,
    stroke,
    strokeWidth,
    width,
    height,
    pixelCoordinates: true,
    lineStartArrow: Boolean(options.lineStartArrow),
    lineEndArrow: Boolean(options.lineEndArrow),
    lineStartArrowType: normalizeShapeArrowType(options.lineStartArrowType || "triangle"),
    lineEndArrowType: normalizeShapeArrowType(options.lineEndArrowType || "triangle"),
    lineStartArrowSize: normalizeShapeArrowSize(options.lineStartArrowSize),
    lineEndArrowSize: normalizeShapeArrowSize(options.lineEndArrowSize),
  };
  let body = "";
  switch (type) {
    case "elbowConnector":
      body = shapePolylineSvgWithArrows(elbowConnectorRoutePointsPx(p1, p2, options.adjust, width, height), arrowOptions);
      break;
    case "curvedConnector":
      {
        const curveOffset = roundSvgNumber((0.5 - (options.adjust?.curve ?? 0.16)) * height * 0.86);
        const c1 = { x: roundSvgNumber(p1.x + (p2.x - p1.x) * 0.38), y: roundSvgNumber(p1.y + curveOffset) };
        const c2 = { x: roundSvgNumber(p1.x + (p2.x - p1.x) * 0.62), y: roundSvgNumber(p2.y + curveOffset) };
        body = shapeBezierSvgWithArrows(p1, c1, c2, p2, arrowOptions);
      }
      break;
    case "freeform":
      body = shapeFreeformSvgWithArrowsPx(p1, p2, width, height, arrowOptions);
      break;
    case "line":
    case "arrow":
    case "doubleArrow":
    default:
      body = shapeLineSvgWithArrows([p1, p2], arrowOptions);
      break;
  }
  body = shapeCompoundLineSvgBody(body, { lineCompound: options.lineCompound, stroke, strokeWidth, fill: "none", isLine: true });
  const effects = shapeEffectsSvg(options, width, height);
  const renderedBody = effects.filterAttr ? `<g ${effects.filterAttr}>${body}</g>` : body;
  const reflectionBody = effects.reflection ? shapeReflectionSvgMarkup(body, effects.reflection) : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" overflow="visible">${effects.defs}${reflectionBody}${renderedBody}</svg>`;
}



function parseSvgViewBoxAttr(viewBox) {
  const parts = String(viewBox || "0 0 100 60")
    .trim()
    .split(/\s+/)
    .map(Number);
  if (parts.length !== 4 || parts.some((value) => !Number.isFinite(value))) {
    return { x: 0, y: 0, width: 100, height: 60 };
  }
  const [x, y, width, height] = parts;
  return {
    x,
    y,
    width: Math.max(1, width),
    height: Math.max(1, height),
  };
}



function shapeBodyFittedToViewBox(body, viewBox, svgWidth, svgHeight, stroke, strokeWidth, lineCompound = "single") {
  if (!body || stroke === "none") return body;
  const box = parseSvgViewBoxAttr(viewBox);
  const widthPx = Math.max(1, Number(svgWidth) || 1);
  const heightPx = Math.max(1, Number(svgHeight) || 1);
  const strokePx = shapeEffectiveOutlineWidth(strokeWidth, lineCompound);
  const insetPx = Math.max(0, strokePx / 2);
  if (!insetPx) return body;
  const padX = Math.min(box.width * 0.45, insetPx * (box.width / widthPx));
  const padY = Math.min(box.height * 0.45, insetPx * (box.height / heightPx));
  if (!(padX > 0 || padY > 0)) return body;
  const scaleX = Math.max(0.01, (box.width - padX * 2) / box.width);
  const scaleY = Math.max(0.01, (box.height - padY * 2) / box.height);
  const transform = [
    `translate(${roundSvgTransformNumber(box.x + padX)} ${roundSvgTransformNumber(box.y + padY)})`,
    `scale(${roundSvgTransformNumber(scaleX)} ${roundSvgTransformNumber(scaleY)})`,
    `translate(${roundSvgTransformNumber(-box.x)} ${roundSvgTransformNumber(-box.y)})`,
  ].join(" ");
  return `<g transform="${transform}">${body}</g>`;
}



function roundSvgTransformNumber(value) {
  return Math.round(Number(value) * 10000) / 10000;
}



function shapeEffectiveOutlineWidth(strokeWidth, lineCompound = "single") {
  const base = normalizeShapeOutlineWidth(strokeWidth);
  switch (normalizeShapeLineCompound(lineCompound)) {
    case "double":
      return base * 2.8;
    case "triple":
      return base * 3.8;
    case "thickThin":
    case "thinThick":
      return base * 3.3;
    default:
      return base;
  }
}



function linePointToPixelPoint(x, y, width, height) {
  return {
    x: roundSvgNumber(clamp01(Number(x)) * Math.max(1, width)),
    y: roundSvgNumber(clamp01(Number(y)) * Math.max(1, height)),
  };
}



function elbowConnectorRoutePointsPx(p1, p2, adjust = {}, width = 180, height = 70) {
  const bend = clampNumber(adjust?.bend, 0, 1, 0.5);
  const dx = Math.abs(p2.x - p1.x);
  const dy = Math.abs(p2.y - p1.y);
  if (adjust?.axis === "y" || (adjust?.axis !== "x" && dy > dx * 1.2)) {
    const bendY = roundSvgNumber(clampNumber(bend * height, Math.min(p1.y, p2.y), Math.max(p1.y, p2.y), (p1.y + p2.y) / 2));
    return [p1, { x: p1.x, y: bendY }, { x: p2.x, y: bendY }, p2];
  }
  const bendX = roundSvgNumber(clampNumber(bend * width, Math.min(p1.x, p2.x), Math.max(p1.x, p2.x), (p1.x + p2.x) / 2));
  return [p1, { x: bendX, y: p1.y }, { x: bendX, y: p2.y }, p2];
}



function shapeFreeformSvgWithArrowsPx(p1, p2, width, height, options = {}) {
  const point = (x, y) => ({ x: roundSvgNumber(width * x), y: roundSvgNumber(height * y) });
  const controls = [
    p1,
    point(0.18, 0.3),
    point(0.32, 0.73),
    point(0.42, 0.4),
    point(0.62, 0.33),
    point(0.7, 0.63),
    point(0.84, 0.78),
    p2,
  ];
  const visible = shapeVisibleLinePoints(controls, options);
  const start = visible[0] || p1;
  const end = visible[visible.length - 1] || p2;
  const path = `<path d="M${start.x} ${start.y} C${controls[1].x} ${controls[1].y} ${controls[2].x} ${controls[2].y} ${controls[3].x} ${controls[3].y} S${controls[4].x} ${controls[4].y} ${controls[5].x} ${controls[5].y} S${controls[6].x} ${controls[6].y} ${end.x} ${end.y}" ${options.lineAttrs}/>`;
  return `${path}${shapeLineArrowHeadsSvg(controls, options)}`;
}



function shapeCompoundLineSvgBody(body, options = {}) {
  const compound = normalizeShapeLineCompound(options.lineCompound);
  const stroke = options.stroke === "none" ? "none" : cssColorForSvg(options.stroke) || "#4472c4";
  if (compound === "single" || stroke === "none") return body;
  const strokeWidth = normalizeShapeOutlineWidth(options.strokeWidth);
  const defs = body.match(/<defs>[\s\S]*?<\/defs>/g)?.join("") || "";
  const rawBody = body.replace(/<defs>[\s\S]*?<\/defs>/g, "");
  const fillOnly = shapeSvgStrokeLayer(rawBody, { stroke: "none", strokeWidth: 0, keepFill: true });
  const gap = shapeCompoundGapColor(options.fill);
  const layers = [fillOnly];
  if (compound === "double") {
    layers.push(
      shapeSvgStrokeLayer(rawBody, { stroke, strokeWidth: strokeWidth * 2.8 }),
      shapeSvgStrokeLayer(rawBody, { stroke: gap, strokeWidth: strokeWidth * 1.2 }),
    );
  } else if (compound === "triple") {
    layers.push(
      shapeSvgStrokeLayer(rawBody, { stroke, strokeWidth: strokeWidth * 3.8 }),
      shapeSvgStrokeLayer(rawBody, { stroke: gap, strokeWidth: strokeWidth * 2.4 }),
      shapeSvgStrokeLayer(rawBody, { stroke, strokeWidth: Math.max(0.6, strokeWidth * 0.8) }),
    );
  } else {
    layers.push(
      shapeSvgStrokeLayer(rawBody, { stroke, strokeWidth: strokeWidth * 3.3 }),
      shapeSvgStrokeLayer(rawBody, { stroke: gap, strokeWidth: strokeWidth * 1.8 }),
      shapeSvgStrokeLayer(rawBody, { stroke, strokeWidth: compound === "thickThin" ? Math.max(0.6, strokeWidth * 0.7) : Math.max(1, strokeWidth * 1.1) }),
    );
  }
  if (options.isLine) layers.push(rawBody);
  return `${defs}${layers.join("")}`;
}



function shapeSvgStrokeLayer(svg, options = {}) {
  let output = svg.replace(/\smarker-(?:start|end)="[^"]*"/g, "");
  if (!options.keepFill) output = output.replace(/\sfill="[^"]*"/g, ' fill="none"');
  output = output.replace(/\sstroke="[^"]*"/g, ` stroke="${escapeXml(options.stroke || "none")}"`);
  output = output.replace(/\sstroke-width="[^"]*"/g, ` stroke-width="${roundSvgNumber(Number(options.strokeWidth) || 0)}"`);
  return output;
}



function shapeCompoundGapColor(fill) {
  if (fill && fill !== "none" && !String(fill).startsWith("url(")) return cssColorForSvg(fill) || "#ffffff";
  return "#ffffff";
}



function shapeLineSvgWithArrows(points, options = {}) {
  const [p1, p2] = shapeVisibleLinePoints(points, options);
  return `<line x1="${p1.x}" y1="${p1.y}" x2="${p2.x}" y2="${p2.y}" ${options.lineAttrs}/>${shapeLineArrowHeadsSvg(points, options)}`;
}



function shapePolylineSvgWithArrows(points, options = {}) {
  const visiblePoints = shapeVisibleLinePoints(points, options);
  return `<polyline points="${visiblePoints.map((point) => `${point.x},${point.y}`).join(" ")}" ${options.lineAttrs}/>${shapeLineArrowHeadsSvg(points, options)}`;
}



function shapeBezierSvgWithArrows(p1, c1, c2, p2, options = {}) {
  const [start, _c1, _c2, end] = shapeVisibleLinePoints([p1, c1, c2, p2], options);
  const path = `<path d="M${start.x} ${start.y} C${c1.x} ${c1.y}, ${c2.x} ${c2.y}, ${end.x} ${end.y}" ${options.lineAttrs}/>`;
  return `${path}${shapeLineArrowHeadsSvg([p1, c1, c2, p2], options)}`;
}



function shapeFreeformSvgWithArrows(p1, p2, options = {}) {
  const controls = [
    p1,
    { x: 18, y: 18 },
    { x: 32, y: 44 },
    { x: 42, y: 24 },
    { x: 62, y: 20 },
    { x: 70, y: 38 },
    { x: 84, y: 47 },
    p2,
  ];
  const visible = shapeVisibleLinePoints(controls, options);
  const start = visible[0] || p1;
  const end = visible[visible.length - 1] || p2;
  const path = `<path d="M${start.x} ${start.y} C18 18 32 44 42 24 S62 20 70 38 S84 47 ${end.x} ${end.y}" ${options.lineAttrs}/>`;
  return `${path}${shapeLineArrowHeadsSvg(controls, options)}`;
}



function shapeVisibleLinePoints(points, options = {}) {
  const visible = (points || []).map((point) => ({ ...point }));
  if (visible.length < 2 || options.stroke === "none") return visible;
  if (options.lineStartArrow && normalizeShapeArrowType(options.lineStartArrowType) !== "open") {
    const segment = firstNonZeroLineSegment(visible, true);
    const endpoint = segment
      ? shapeArrowLineEndpoint(segment.to, segment.from, options.lineStartArrowType, options.strokeWidth, options.lineStartArrowSize, options)
      : null;
    if (endpoint) visible[0] = endpoint;
  }
  if (options.lineEndArrow && normalizeShapeArrowType(options.lineEndArrowType) !== "open") {
    const segment = firstNonZeroLineSegment(visible, false);
    const endpoint = segment
      ? shapeArrowLineEndpoint(segment.to, segment.from, options.lineEndArrowType, options.strokeWidth, options.lineEndArrowSize, options)
      : null;
    if (endpoint) visible[visible.length - 1] = endpoint;
  }
  return visible;
}



function shapeArrowLineEndpoint(tip, from, type, strokeWidth, size, options = {}) {
  const geometry = shapeArrowGeometry(tip, from, strokeWidth, size, options);
  if (!geometry) return null;
  if (normalizeShapeArrowType(type) === "diamond" || normalizeShapeArrowType(type) === "oval") return geometry.center || geometry.back;
  return geometry.back;
}



function shapeLineArrowHeadsSvg(points, options = {}) {
  if (options.stroke === "none") return "";
  const parts = [];
  if (options.lineStartArrow) {
    const segment = firstNonZeroLineSegment(points, true);
    if (segment) parts.push(shapeArrowHeadSvg(segment.to, segment.from, options.lineStartArrowType, options.stroke, options.strokeWidth, options.lineStartArrowSize, options));
  }
  if (options.lineEndArrow) {
    const segment = firstNonZeroLineSegment(points, false);
    if (segment) parts.push(shapeArrowHeadSvg(segment.to, segment.from, options.lineEndArrowType, options.stroke, options.strokeWidth, options.lineEndArrowSize, options));
  }
  return parts.join("");
}



function firstNonZeroLineSegment(points, startSide) {
  const list = (points || []).filter(Boolean);
  if (list.length < 2) return null;
  const indexes = startSide ? [...Array(list.length - 1).keys()] : [...Array(list.length - 1).keys()].reverse();
  for (const index of indexes) {
    const from = startSide ? list[index + 1] : list[index];
    const to = startSide ? list[index] : list[index + 1];
    if (Math.abs(to.x - from.x) > 0.01 || Math.abs(to.y - from.y) > 0.01) return { from, to };
  }
  return null;
}



function shapeArrowHeadSvg(tip, from, type, color, strokeWidth, size, options = {}) {
  const screen = shapeArrowGeometry(tip, from, strokeWidth, size, options);
  if (!screen) return "";
  const stroke = escapeXml(color);
  const fill = type === "open" ? "none" : stroke;
  const toSvg = (point) => `${roundSvgNumber(point.x)},${roundSvgNumber(point.y)}`;
  const normalizedType = normalizeShapeArrowType(type) || "triangle";
  if (normalizedType === "open") {
    return `<path d="M${toSvg(screen.left)} L${toSvg(screen.tip)} L${toSvg(screen.right)}" fill="none" stroke="${stroke}" stroke-width="${roundSvgNumber(Math.max(1.1, Number(strokeWidth) || 1.2))}" stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke"/>`;
  }
  if (normalizedType === "diamond") {
    return `<polygon points="${[screen.tip, screen.left, screen.back, screen.right].map(toSvg).join(" ")}" fill="${fill}" stroke="${stroke}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
  }
  if (normalizedType === "oval") {
    const rx = roundSvgNumber(screen.radiusX);
    const ry = roundSvgNumber(screen.radiusY);
    return `<ellipse cx="${roundSvgNumber(screen.center.x)}" cy="${roundSvgNumber(screen.center.y)}" rx="${rx}" ry="${ry}" fill="${fill}" stroke="${stroke}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
  }
  if (normalizedType === "stealth") {
    return `<polygon points="${[screen.tip, screen.left, screen.notch, screen.right].map(toSvg).join(" ")}" fill="${fill}" stroke="${stroke}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
  }
  return `<polygon points="${[screen.tip, screen.left, screen.right].map(toSvg).join(" ")}" fill="${fill}" stroke="${stroke}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
}



function shapeArrowGeometry(tip, from, strokeWidth, size, options = {}) {
  const sx = options.pixelCoordinates ? 1 : (Number(options.width) || 150) / 100;
  const sy = options.pixelCoordinates ? 1 : (Number(options.height) || 100) / 60;
  const tipScreen = { x: tip.x * sx, y: tip.y * sy };
  const fromScreen = { x: from.x * sx, y: from.y * sy };
  const dx = tipScreen.x - fromScreen.x;
  const dy = tipScreen.y - fromScreen.y;
  const length = Math.hypot(dx, dy);
  if (!Number.isFinite(length) || length < 0.01) return null;
  const ux = dx / length;
  const uy = dy / length;
  const px = -uy;
  const py = ux;
  const base = shapeArrowBaseLength(strokeWidth, size);
  const half = shapeArrowHalfWidth(strokeWidth, size);
  const back = { x: tipScreen.x - ux * base, y: tipScreen.y - uy * base };
  const left = { x: back.x + px * half, y: back.y + py * half };
  const right = { x: back.x - px * half, y: back.y - py * half };
  const notch = { x: tipScreen.x - ux * base * 0.62, y: tipScreen.y - uy * base * 0.62 };
  const center = { x: tipScreen.x - ux * base * 0.48, y: tipScreen.y - uy * base * 0.48 };
  const toView = (point) => ({ x: point.x / sx, y: point.y / sy });
  return {
    tip: toView(tipScreen),
    left: toView(left),
    right: toView(right),
    back: toView(back),
    notch: toView(notch),
    center: toView(center),
    radiusX: Math.max(1.5, half / sx),
    radiusY: Math.max(1.5, half / sy),
  };
}



function shapeArrowBaseLength(strokeWidth, size) {
  const scale = { small: 0.76, medium: 1, large: 1.32 }[normalizeShapeArrowSize(size)] || 1;
  return Math.max(6.2, (Number(strokeWidth) || DEFAULT_SHAPE_OUTLINE_WIDTH) * 2.4 + 5.4) * scale;
}



function shapeArrowHalfWidth(strokeWidth, size) {
  return shapeArrowBaseLength(strokeWidth, size) * 0.36;
}



function shapeArrowMarkerSvg(id, type, color, isStart = false, size = "medium") {
  const stroke = escapeXml(color);
  const fill = type === "open" ? "none" : stroke;
  const scale = { small: 0.75, medium: 1, large: 1.35 }[normalizeShapeArrowSize(size)] || 1;
  const markerWidth = roundSvgNumber(10 * scale);
  const markerHeight = roundSvgNumber(10 * scale);
  const common = `id="${escapeAttr(id)}" markerWidth="${markerWidth}" markerHeight="${markerHeight}" viewBox="0 0 10 10" refX="8" refY="5" orient="${isStart ? "auto-start-reverse" : "auto"}" markerUnits="strokeWidth"`;
  switch (normalizeShapeArrowType(type) || "triangle") {
    case "stealth":
      return `<marker ${common}><path d="M0 0 L10 5 L0 10 L3.5 5 Z" fill="${fill}" stroke="${stroke}" stroke-width="1"/></marker>`;
    case "diamond":
      return `<marker ${common} refX="9"><path d="M0 5 L5 0 L10 5 L5 10 Z" fill="${fill}" stroke="${stroke}" stroke-width="1"/></marker>`;
    case "oval":
      return `<marker ${common} refX="9"><circle cx="5" cy="5" r="4" fill="${fill}" stroke="${stroke}" stroke-width="1"/></marker>`;
    case "open":
      return `<marker ${common}><path d="M1 1 L9 5 L1 9" fill="none" stroke="${stroke}" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></marker>`;
    case "triangle":
    default:
      return `<marker ${common}><path d="M0 0 L10 5 L0 10 Z" fill="${fill}" stroke="${stroke}" stroke-width="1"/></marker>`;
  }
}



function cssColorForSvg(value) {
  const raw = String(value || "").trim();
  if (/^#[0-9a-f]{6}$/i.test(raw)) return raw;
  if (/^rgba?\(\s*\d+\s*,\s*\d+\s*,\s*\d+(?:\s*,\s*(?:0|1|0?\.\d+))?\s*\)$/i.test(raw)) return raw;
  return "";
}



function shapeGradientDefSvg(gradient) {
  const vector = shapeGradientVector(gradient.angle);
  const stops = gradient.stops
    .map((stop) => {
      const color = svgColorParts(stop.color);
      return `<stop offset="${roundSvgNumber(stop.pos * 100)}%" stop-color="${escapeXml(color.color)}"${color.opacity < 1 ? ` stop-opacity="${roundSvgNumber(color.opacity)}"` : ""}/>`;
    })
    .join("");
  return `<defs><linearGradient id="shapeFillGradient" gradientUnits="objectBoundingBox" x1="${vector.x1}" y1="${vector.y1}" x2="${vector.x2}" y2="${vector.y2}">${stops}</linearGradient></defs>`;
}



function shapePatternDefSvg(pattern) {
  const foreground = svgColorParts(pattern.foreground);
  const background = svgColorParts(pattern.background);
  const bg = `<rect width="8" height="8" fill="${escapeXml(background.color)}"${background.opacity < 1 ? ` opacity="${roundSvgNumber(background.opacity)}"` : ""}/>`;
  const fgAttr = `stroke="${escapeXml(foreground.color)}"${foreground.opacity < 1 ? ` opacity="${roundSvgNumber(foreground.opacity)}"` : ""} stroke-width="1"`;
  const fillDot = `<circle cx="2" cy="2" r="0.7" fill="${escapeXml(foreground.color)}"${foreground.opacity < 1 ? ` opacity="${roundSvgNumber(foreground.opacity)}"` : ""}/>`;
  const body = {
    pct5: `${bg}${fillDot}`,
    pct10: `${bg}${fillDot}<circle cx="6" cy="6" r="0.7" fill="${escapeXml(foreground.color)}"${foreground.opacity < 1 ? ` opacity="${roundSvgNumber(foreground.opacity)}"` : ""}/>`,
    pct20: `${bg}<path d="M0 0 L8 8 M8 0 L0 8" ${fgAttr}/>`,
    pct25: `${bg}<path d="M0 0 H8 M0 4 H8" ${fgAttr}/>`,
    pct50: `${bg}<path d="M0 0 H8 M0 2 H8 M0 4 H8 M0 6 H8" ${fgAttr}/>`,
    ltHorz: `${bg}<path d="M0 4 H8" ${fgAttr}/>`,
    ltVert: `${bg}<path d="M4 0 V8" ${fgAttr}/>`,
    ltDnDiag: `${bg}<path d="M-2 0 L8 10 M2 -2 L10 6" ${fgAttr}/>`,
    ltUpDiag: `${bg}<path d="M-2 8 L8 -2 M2 10 L10 2" ${fgAttr}/>`,
    diagCross: `${bg}<path d="M-2 0 L8 10 M2 -2 L10 6 M-2 8 L8 -2 M2 10 L10 2" ${fgAttr}/>`,
    smGrid: `${bg}<path d="M4 0 V8 M0 4 H8" ${fgAttr}/>`,
  }[pattern.preset] || `${bg}<path d="M0 0 L8 8 M8 0 L0 8" ${fgAttr}/>`;
  return `<defs><pattern id="shapeFillPattern" patternUnits="userSpaceOnUse" width="8" height="8">${body}</pattern></defs>`;
}



function shapeEffectsSvg(options = {}, width = 150, height = 100) {
  const shadow = shapeShadowRenderModel(options.shadow);
  const glow = normalizeShapeGlow(options.glow);
  const softEdges = normalizeShapeSoftEdges(options.softEdges);
  const reflection = normalizeShapeReflection(options.reflection);
  const primitives = [];
  if (shadow) {
    const color = svgColorParts(shadow.color);
    primitives.push(`<feDropShadow dx="${shapeEffectX(shadow.dx, width)}" dy="${shapeEffectY(shadow.dy, height)}" stdDeviation="${shapeEffectRadius(shadow.blur, width, height)}" flood-color="${escapeXml(color.color)}" flood-opacity="${roundSvgNumber(color.opacity)}"/>`);
  }
  if (glow) {
    const color = svgColorParts(glow.color);
    primitives.push(`<feDropShadow dx="0" dy="0" stdDeviation="${shapeEffectRadius(glow.radius, width, height)}" flood-color="${escapeXml(color.color)}" flood-opacity="${roundSvgNumber(Math.max(0.35, color.opacity))}"/>`);
  }
  if (softEdges) {
    primitives.push(`<feGaussianBlur stdDeviation="${shapeEffectRadius(softEdges, width, height)}"/>`);
  }
  return {
    defs: primitives.length ? `<defs><filter id="shapeEffects" x="-35%" y="-35%" width="170%" height="190%" color-interpolation-filters="sRGB">${primitives.join("")}</filter></defs>` : "",
    filterAttr: primitives.length ? 'filter="url(#shapeEffects)"' : "",
    reflection,
  };
}



function shapeShadowRenderModel(shadow) {
  const normalized = normalizeShapeShadow(shadow);
  if (!normalized) return null;
  if (typeof normalized === "object") return normalized;
  if (normalized === "outer2") return { type: "outer", dx: 0, dy: 3, blur: 5, color: "rgba(0,0,0,0.24)" };
  if (normalized === "outer1") return { type: "outer", dx: 2, dy: 2, blur: 2.5, color: "rgba(0,0,0,0.28)" };
  return null;
}



function shapeEffectX(value, width) {
  return roundSvgNumber((Number(value) || 0) * 100 / Math.max(1, Number(width) || 150));
}



function shapeEffectY(value, height) {
  return roundSvgNumber((Number(value) || 0) * 60 / Math.max(1, Number(height) || 100));
}



function shapeEffectRadius(value, width, height) {
  const scale = (100 / Math.max(1, Number(width) || 150) + 60 / Math.max(1, Number(height) || 100)) / 2;
  return roundSvgNumber(Math.max(0.1, (Number(value) || 0.8) * scale));
}



function shapeReflectionSvgMarkup(body, reflection) {
  if (!reflection) return "";
  const preset = reflection.preset || "tight";
  const opacity = roundSvgNumber(Number(reflection.alpha) || ({ tight: 0.22, half: 0.34, full: 0.48 }[preset] || 0.26));
  const scale = preset === "full" ? 0.48 : preset === "half" ? 0.34 : 0.22;
  const translate = roundSvgNumber(64 + scale * 60);
  const blur = shapeEffectRadius(reflection.blur || 1.2, 150, 100);
  const filter = blur ? `<defs><filter id="shapeReflectionBlur" x="-20%" y="-20%" width="140%" height="160%"><feGaussianBlur stdDeviation="${blur}"/></filter></defs>` : "";
  return `${filter}<g opacity="${opacity}" transform="translate(0 ${translate}) scale(1 -${roundSvgNumber(scale)})" filter="url(#shapeReflectionBlur)">${body}</g>`;
}



function svgColorParts(value) {
  const raw = String(value || "").trim();
  const rgba = raw.match(/^rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(0|1|0?\.\d+)\s*\)$/i);
  if (rgba) {
    const rgb = [rgba[1], rgba[2], rgba[3]].map((part) => Math.max(0, Math.min(255, Number(part) || 0)));
    return { color: rgbToHex(rgb), opacity: Math.max(0, Math.min(1, Number(rgba[4]))) };
  }
  return { color: cssColorForSvg(raw) || "#000000", opacity: 1 };
}



function shapeGradientVector(angle = 0) {
  const radians = (normalizeRotationDegrees(angle) * Math.PI) / 180;
  const dx = Math.cos(radians) / 2;
  const dy = -Math.sin(radians) / 2;
  return {
    x1: roundSvgNumber(0.5 - dx),
    y1: roundSvgNumber(0.5 - dy),
    x2: roundSvgNumber(0.5 + dx),
    y2: roundSvgNumber(0.5 + dy),
  };
}



function shapeLineDashArray(type, width = DEFAULT_SHAPE_OUTLINE_WIDTH) {
  const unit = Math.max(1, Number(width) || DEFAULT_SHAPE_OUTLINE_WIDTH);
  switch (normalizeShapeLineDash(type)) {
    case "shortDash":
      return `${roundSvgNumber(unit * 2.5)} ${roundSvgNumber(unit * 2)}`;
    case "dash":
      return `${roundSvgNumber(unit * 4)} ${roundSvgNumber(unit * 3)}`;
    case "dot":
      return `${roundSvgNumber(unit)} ${roundSvgNumber(unit * 2.2)}`;
    case "squareDot":
      return `${roundSvgNumber(unit * 1.4)} ${roundSvgNumber(unit * 2.4)}`;
    case "dashDot":
      return `${roundSvgNumber(unit * 4)} ${roundSvgNumber(unit * 2)} ${roundSvgNumber(unit)} ${roundSvgNumber(unit * 2)}`;
    case "longDash":
      return `${roundSvgNumber(unit * 7)} ${roundSvgNumber(unit * 3)}`;
    case "longDashDot":
      return `${roundSvgNumber(unit * 7)} ${roundSvgNumber(unit * 3)} ${roundSvgNumber(unit)} ${roundSvgNumber(unit * 3)}`;
    case "longDashDotDot":
      return `${roundSvgNumber(unit * 7)} ${roundSvgNumber(unit * 2.5)} ${roundSvgNumber(unit)} ${roundSvgNumber(unit * 2.5)} ${roundSvgNumber(unit)} ${roundSvgNumber(unit * 2.5)}`;
    default:
      return "";
  }
}



function elbowConnectorRoutePoints(p1, p2, adjust = {}) {
  const bend = clampNumber(adjust?.bend, 0, 1, 0.5);
  const dx = Math.abs(p2.x - p1.x);
  const dy = Math.abs(p2.y - p1.y);
  if (adjust?.axis === "y" || (adjust?.axis !== "x" && dy > dx * 1.2)) {
    const bendY = roundSvgNumber(clampNumber(6 + bend * 48, Math.min(p1.y, p2.y), Math.max(p1.y, p2.y), (p1.y + p2.y) / 2));
    return [p1, { x: p1.x, y: bendY }, { x: p2.x, y: bendY }, p2];
  }
  const bendX = roundSvgNumber(clampNumber(6 + bend * 88, Math.min(p1.x, p2.x), Math.max(p1.x, p2.x), (p1.x + p2.x) / 2));
  return [p1, { x: bendX, y: p1.y }, { x: bendX, y: p2.y }, p2];
}



function shapeCustomPathSvgD(customPath) {
  return normalizeShapeCustomPath(customPath)?.commands
    .map((command) => {
      if (command.cmd === "Z") return "Z";
      const points = command.points.map((point) => `${roundSvgNumber(point.x * 100)} ${roundSvgNumber(point.y * 60)}`);
      return `${command.cmd}${points.join(" ")}`;
    })
    .join(" ") || "M0 0 L100 0 L100 60 L0 60 Z";
}



function shapeTextSvgMarkup(text, options = {}) {
  const width = Math.max(1, Number(options.width) || 150);
  const height = Math.max(1, Number(options.height) || 100);
  const layer = shapeTextLayerHtml({shape:{...options,text}}, 1);
  if (!layer) return "";
  // The exact live HTML layout also renders exported SVG/Excel images. No
  // character-count wrapping, alternate font scale or line truncation.
  const markup=layer.replace(/<br>/g,"<br/>");
  return `<foreignObject width="${width}" height="${height}" transform="scale(${100/width} ${60/height})"><div xmlns="http://www.w3.org/1999/xhtml" style="position:relative;width:${width}px;height:${height}px">${markup}</div></foreignObject>`;
}



function mathShapeText(type) {
  const texts = {
    mathPlus: "+",
    mathMinus: "-",
    mathMultiply: "×",
    mathDivide: "÷",
    mathEqual: "=",
    mathNotEqual: "≠",
  };
  return texts[type] || "+";
}



function regularPolygonPoints(cx, cy, radius, count, startDegrees = -90, options = {}) {
  const points = Array.from({ length: count }, (_value, index) => {
    const angle = ((startDegrees + (360 / count) * index) * Math.PI) / 180;
    return [cx + Math.cos(angle) * radius, cy + Math.sin(angle) * radius];
  });
  return options.fit ? fitSvgPointsString(points) : svgPointsString(points);
}



function starPoints(cx, cy, outerRadius, innerRadius, points, options = {}) {
  const star = Array.from({ length: points * 2 }, (_value, index) => {
    const angle = ((-90 + (180 / points) * index) * Math.PI) / 180;
    const radius = index % 2 === 0 ? outerRadius : innerRadius;
    return [cx + Math.cos(angle) * radius, cy + Math.sin(angle) * radius];
  });
  return options.fit ? fitSvgPointsString(star) : svgPointsString(star);
}



function star4Points(inner = 0.62) {
  const normalized = clampNumber(inner, 0.18, 0.82, 0.62);
  const flatStart = roundSvgNumber(50 - normalized * 28);
  const flatEnd = roundSvgNumber(50 + normalized * 28);
  const notchLeft = roundSvgNumber(50 - normalized * 48);
  const notchRight = roundSvgNumber(50 + normalized * 48);
  return svgPointsString([
    [0.2, 15],
    [flatStart, 15],
    [50, 0.2],
    [flatEnd, 15],
    [99.8, 15],
    [notchRight, 30],
    [99.8, 45],
    [flatEnd, 45],
    [50, 59.8],
    [flatStart, 45],
    [0.2, 45],
    [notchLeft, 30],
  ]);
}



function explosionSealPoints() {
  return officeDrawingPointsString([
    [10800, 5800],
    [8352, 2295],
    [7312, 6320],
    [370, 2295],
    [4627, 7617],
    [0, 8615],
    [3722, 11775],
    [135, 14587],
    [5667, 13937],
    [4762, 17617],
    [7715, 15627],
    [8485, 21600],
    [10532, 14935],
    [13247, 19737],
    [14020, 14457],
    [18145, 18095],
    [16837, 12942],
    [21600, 13290],
    [17607, 10475],
    [21097, 8137],
    [16702, 7315],
    [18380, 4457],
    [14155, 5325],
    [14522, 0],
  ]);
}



function explosionBurstPoints() {
  return officeDrawingPointsString([
    [11462, 4342],
    [9722, 1887],
    [8550, 6382],
    [4502, 3625],
    [5372, 7817],
    [1172, 8270],
    [3935, 11592],
    [0, 12877],
    [3330, 15370],
    [1285, 17825],
    [4805, 18240],
    [4917, 21600],
    [7527, 18125],
    [8700, 19712],
    [9872, 17370],
    [11612, 18842],
    [12180, 15935],
    [14942, 17370],
    [14640, 14350],
    [18877, 15632],
    [16380, 12310],
    [18270, 11290],
    [16985, 9402],
    [21600, 6645],
    [16380, 6532],
    [18007, 3172],
    [14525, 5777],
    [14790, 0],
  ]);
}



function officeDrawingPointsString(points) {
  return svgPointsString(points.map(([x, y]) => [(x / 21600) * 100, (y / 21600) * 60]));
}



function isCalloutShapeType(type) {
  return ["calloutRect", "calloutRoundRect", "calloutOval", "calloutCloud"].includes(type);
}



function calloutGeometry(adjust = {}, type = "calloutRect") {
  const rawTail = {
    x: clampNumber(adjust.tailX, -2, 3, defaultCalloutTailX(type)),
    y: clampNumber(adjust.tailY, -2, 3, defaultCalloutTailY(type)),
  };
  const tail = calloutRenderableTailPoint(type, rawTail);
  const tailX = tail.x;
  const tailY = tail.y;
  const body = { left: 0, top: 0, right: 100, bottom: 60 };
  const tip = { x: tailX * 100, y: tailY * 60 };
  const office = calloutWedgeOfficeGeometry(tailX, tailY);
  const viewBox = calloutViewBoxFromPoints([tip, ...office.points, { x: body.left, y: body.top }, { x: body.right, y: body.bottom }]);
  return {
    edge: office.edge,
    body,
    tip: { x: roundSvgNumber(tip.x), y: roundSvgNumber(tip.y) },
    base: office.base,
    points: office.points,
    viewBox,
  };
}



function calloutWedgeOfficeGeometry(tailX, tailY) {
  const w = 100;
  const h = 60;
  const l = 0;
  const t = 0;
  const r = 100;
  const b = 60;
  const hc = 50;
  const vc = 30;
  const xPos = tailX * w;
  const yPos = tailY * h;
  const dxPos = xPos - hc;
  const dyPos = yPos - vc;
  const choose = (test, yes, no) => (test > 0 ? yes : no);
  const dq = (dxPos * h) / w;
  const ady = Math.abs(dyPos);
  const adq = Math.abs(dq);
  const dz = ady - adq;
  const xg1 = choose(dxPos, 7, 2);
  const xg2 = choose(dxPos, 10, 5);
  const x1 = (w * xg1) / 12;
  const x2 = (w * xg2) / 12;
  const yg1 = choose(dyPos, 7, 2);
  const yg2 = choose(dyPos, 10, 5);
  const y1 = (h * yg1) / 12;
  const y2 = (h * yg2) / 12;
  const t1 = choose(dxPos, l, xPos);
  const xl = choose(dz, l, t1);
  const t2 = choose(dyPos, x1, xPos);
  const xt = choose(dz, t2, x1);
  const t3 = choose(dxPos, xPos, r);
  const xr = choose(dz, r, t3);
  const t4 = choose(dyPos, xPos, x1);
  const xb = choose(dz, t4, x1);
  const t5 = choose(dxPos, y1, yPos);
  const yl = choose(dz, y1, t5);
  const t6 = choose(dyPos, t, yPos);
  const yt = choose(dz, t6, t);
  const t7 = choose(dxPos, yPos, y1);
  const yr = choose(dz, y1, t7);
  const t8 = choose(dyPos, yPos, b);
  const yb = choose(dz, t8, b);
  const points = [
    { x: l, y: t },
    { x: x1, y: t },
    { x: xt, y: yt },
    { x: x2, y: t },
    { x: r, y: t },
    { x: r, y: y1 },
    { x: xr, y: yr },
    { x: r, y: y2 },
    { x: r, y: b },
    { x: x2, y: b },
    { x: xb, y: yb },
    { x: x1, y: b },
    { x: l, y: b },
    { x: l, y: y2 },
    { x: xl, y: yl },
    { x: l, y: y1 },
  ].map((point) => ({ x: roundSvgNumber(point.x), y: roundSvgNumber(point.y) }));
  const edge = Math.abs(xb - xPos) < 0.01 && Math.abs(yb - yPos) < 0.01
    ? "bottom"
    : Math.abs(xt - xPos) < 0.01 && Math.abs(yt - yPos) < 0.01
      ? "top"
      : Math.abs(xl - xPos) < 0.01 && Math.abs(yl - yPos) < 0.01
        ? "left"
        : "right";
  const base = {
    bottom: [{ x: x2, y: b }, { x: x1, y: b }],
    top: [{ x: x1, y: t }, { x: x2, y: t }],
    left: [{ x: l, y: y2 }, { x: l, y: y1 }],
    right: [{ x: r, y: y1 }, { x: r, y: y2 }],
  }[edge].map((point) => ({ x: roundSvgNumber(point.x), y: roundSvgNumber(point.y) }));
  return { edge, base, points };
}



function calloutViewBoxFromPoints(points) {
  const xs = points.map((point) => Number(point.x)).filter(Number.isFinite);
  const ys = points.map((point) => Number(point.y)).filter(Number.isFinite);
  const left = Math.min(0, ...xs);
  const top = Math.min(0, ...ys);
  const right = Math.max(100, ...xs);
  const bottom = Math.max(60, ...ys);
  return {
    left: roundSvgNumber(left),
    top: roundSvgNumber(top),
    right: roundSvgNumber(right),
    bottom: roundSvgNumber(bottom),
    width: roundSvgNumber(Math.max(1, right - left)),
    height: roundSvgNumber(Math.max(1, bottom - top)),
  };
}



function calloutViewBoxAttr(adjust = {}, type = "calloutRect") {
  const viewBox = calloutGeometry(adjust, type).viewBox;
  return `${viewBox.left} ${viewBox.top} ${viewBox.width} ${viewBox.height}`;
}



function calloutShapePath(type, adjust = {}, shapeAttrs = "") {
  const geometry = calloutGeometry(adjust, type);
  if (type === "calloutRect") {
    return `<path d="${calloutUnifiedPathD(type, geometry)}" ${shapeAttrs}/>`;
  }
  if (type === "calloutRoundRect") {
    return `<path d="${calloutUnifiedPathD(type, geometry)}" ${shapeAttrs}/>`;
  }
  if (type === "calloutOval") {
    return `<path d="${calloutUnifiedPathD(type, geometry)}" ${shapeAttrs}/>`;
  }
  return `<path d="${calloutUnifiedPathD(type, geometry)}" ${shapeAttrs}/>`;
}



function calloutTailMarkup(geometry, type = "calloutRect", shapeAttrs = "") {
  const fillAttrs = calloutFillOnlyAttrs(shapeAttrs);
  const lineAttrs = calloutStrokeOnlyAttrs(shapeAttrs);
  return `<path d="${calloutTailFillPathD(geometry, type)}" ${fillAttrs}/><path d="${calloutTailOutlinePathD(geometry, type)}" ${lineAttrs}/>`;
}



function calloutTailFillPathD(geometry, type = "calloutRect") {
  const [a, b] = calloutTailBasePoints(geometry, type);
  const { tip } = geometry;
  return `M${a.x} ${a.y} L${tip.x} ${tip.y} L${b.x} ${b.y} Z`;
}



function calloutTailOutlinePathD(geometry, type = "calloutRect") {
  const [a, b] = calloutTailOutlineBasePoints(geometry, type);
  const { tip } = geometry;
  return `M${a.x} ${a.y} L${tip.x} ${tip.y} L${b.x} ${b.y}`;
}



function calloutFillOnlyAttrs(shapeAttrs = "") {
  return String(shapeAttrs)
    .replace(/\sstroke="[^"]*"/, ' stroke="none"')
    .replace(/\sstroke-width="[^"]*"/, ' stroke-width="0"');
}



function calloutStrokeOnlyAttrs(shapeAttrs = "") {
  return String(shapeAttrs).replace(/\sfill="[^"]*"/, ' fill="none"');
}



function calloutTailBasePoints(geometry, type = "calloutRect") {
  const base = geometry.base || [];
  if (type === "calloutRect") return base;
  const [a, b] = base;
  if (!a || !b) return base;
  const { body } = geometry;
  const edge = calloutEffectiveEdge(geometry);
  const overlap = type === "calloutRoundRect" ? 4 : 10;
  const inside = (point) => {
    if (edge === "bottom") return { ...point, y: roundSvgNumber(Math.max(body.top, body.bottom - overlap)) };
    if (edge === "top") return { ...point, y: roundSvgNumber(Math.min(body.bottom, body.top + overlap)) };
    if (edge === "left") return { ...point, x: roundSvgNumber(Math.min(body.right, body.left + overlap)) };
    return { ...point, x: roundSvgNumber(Math.max(body.left, body.right - overlap)) };
  };
  return [inside(a), inside(b)];
}



function calloutTailOutlineBasePoints(geometry, type = "calloutRect") {
  const base = calloutTailBoundaryTargets(geometry, type);
  if (type === "calloutOval") {
    return base.map((point) => calloutEllipseBoundaryPoint(geometry.tip, point, geometry.body) || point);
  }
  if (type === "calloutCloud") {
    return base.map((point) => calloutCloudBoundaryPoint(geometry.tip, point, geometry.body) || calloutEllipseBoundaryPoint(geometry.tip, point, geometry.body) || point);
  }
  return base;
}



function calloutEllipseBoundaryPoint(tip, base, body) {
  const cx = (body.left + body.right) / 2;
  const cy = (body.top + body.bottom) / 2;
  const rx = Math.max(1, (body.right - body.left) / 2);
  const ry = Math.max(1, (body.bottom - body.top) / 2);
  const dx = base.x - tip.x;
  const dy = base.y - tip.y;
  const px = tip.x - cx;
  const py = tip.y - cy;
  const a = (dx * dx) / (rx * rx) + (dy * dy) / (ry * ry);
  const b = (2 * px * dx) / (rx * rx) + (2 * py * dy) / (ry * ry);
  const c = (px * px) / (rx * rx) + (py * py) / (ry * ry) - 1;
  const discriminant = b * b - 4 * a * c;
  if (!Number.isFinite(discriminant) || discriminant < 0 || !a) return null;
  const roots = [
    (-b - Math.sqrt(discriminant)) / (2 * a),
    (-b + Math.sqrt(discriminant)) / (2 * a),
  ].filter((value) => Number.isFinite(value) && value >= 0);
  if (!roots.length) return null;
  const t = Math.min(...roots);
  return {
    x: roundSvgNumber(tip.x + dx * t),
    y: roundSvgNumber(tip.y + dy * t),
  };
}



function calloutCloudBoundaryPoint(tip, base, body) {
  const outline = cloudOutlinePoints(body, 14);
  if (outline.length < 2) return null;
  const rayEnd = calloutBoundaryRayEnd(tip, base);
  let best = null;
  for (let index = 0; index < outline.length; index += 1) {
    const current = outline[index];
    const next = outline[(index + 1) % outline.length];
    const intersection = segmentIntersectionPoint(tip, rayEnd, current, next);
    if (!intersection) continue;
    if (!best || intersection.t < best.t) best = intersection;
  }
  return best ? { x: roundSvgNumber(best.x), y: roundSvgNumber(best.y) } : null;
}



function calloutTailBoundaryTargets(geometry, type = "calloutRect") {
  const base = geometry.base || [];
  const { body, tip } = geometry;
  const edge = calloutEffectiveEdge(geometry);
  const horizontalSpread = type === "calloutCloud" ? 18 : type === "calloutOval" ? 20 : 28;
  const verticalSpread = type === "calloutCloud" ? 12 : type === "calloutOval" ? 14 : 18;
  const targetInset = type === "calloutCloud" ? 10 : type === "calloutOval" ? 8 : 4;
  if (edge === "bottom" || edge === "top") {
    const half = horizontalSpread / 2;
    const center = clampNumber(tip.x, body.left + half, body.right - half, (body.left + body.right) / 2);
    const y = edge === "bottom"
      ? roundSvgNumber(Math.max(body.top, body.bottom - targetInset))
      : roundSvgNumber(Math.min(body.bottom, body.top + targetInset));
    const left = { x: roundSvgNumber(center - half), y };
    const right = { x: roundSvgNumber(center + half), y };
    return edge === "bottom" ? [right, left] : [left, right];
  }
  const half = verticalSpread / 2;
  const center = clampNumber(tip.y, body.top + half, body.bottom - half, (body.top + body.bottom) / 2);
  const x = edge === "left"
    ? roundSvgNumber(Math.min(body.right, body.left + targetInset))
    : roundSvgNumber(Math.max(body.left, body.right - targetInset));
  const top = { x, y: roundSvgNumber(center - half) };
  const bottom = { x, y: roundSvgNumber(center + half) };
  return edge === "left" ? [bottom, top] : [top, bottom];
}



function calloutEffectiveEdge(geometry) {
  const { body, tip } = geometry;
  const candidates = [
    { edge: "bottom", amount: tip.y - body.bottom },
    { edge: "top", amount: body.top - tip.y },
    { edge: "right", amount: tip.x - body.right },
    { edge: "left", amount: body.left - tip.x },
  ].filter((item) => Number.isFinite(item.amount) && item.amount > 0);
  if (candidates.length) {
    candidates.sort((a, b) => b.amount - a.amount);
    return candidates[0].edge;
  }
  const distances = [
    { edge: "bottom", distance: Math.abs(body.bottom - tip.y) },
    { edge: "top", distance: Math.abs(tip.y - body.top) },
    { edge: "right", distance: Math.abs(body.right - tip.x) },
    { edge: "left", distance: Math.abs(tip.x - body.left) },
  ];
  distances.sort((a, b) => a.distance - b.distance);
  if (Number.isFinite(distances[0]?.distance)) return distances[0].edge;
  return geometry.edge || "bottom";
}



function calloutRenderableTailPoint(type, point) {
  const normalized = {
    x: clampNumber(point?.x, -2, 3, defaultCalloutTailX(type)),
    y: clampNumber(point?.y, -2, 3, defaultCalloutTailY(type)),
  };
  if (!isCalloutShapeType(type)) return normalized;
  if (normalized.x < 0 || normalized.x > 1 || normalized.y < 0 || normalized.y > 1) return normalized;
  const guard = 0.045;
  const epsilon = type === "calloutCloud" ? 0.026 : type === "calloutOval" ? 0.022 : 0.018;
  const nearTop = normalized.y <= guard;
  const nearBottom = normalized.y >= 1 - guard;
  const nearLeft = normalized.x <= guard;
  const nearRight = normalized.x >= 1 - guard;
  const nearHorizontalEdge = nearTop || nearBottom;
  const nearVerticalEdge = nearLeft || nearRight;
  const nearCorner = nearHorizontalEdge && nearVerticalEdge;
  if (!nearHorizontalEdge && !nearVerticalEdge) return normalized;
  const next = { ...normalized };
  if (nearCorner) {
    if (nearLeft) next.x = -epsilon;
    if (nearRight) next.x = 1 + epsilon;
    if (nearTop) next.y = -epsilon;
    if (nearBottom) next.y = 1 + epsilon;
    return {
      x: clampNumber(next.x, -2, 3, defaultCalloutTailX(type)),
      y: clampNumber(next.y, -2, 3, defaultCalloutTailY(type)),
    };
  }
  const distances = [
    { edge: "top", distance: normalized.y },
    { edge: "bottom", distance: 1 - normalized.y },
    { edge: "left", distance: normalized.x },
    { edge: "right", distance: 1 - normalized.x },
  ].sort((a, b) => a.distance - b.distance);
  const edge = distances[0]?.edge;
  if (edge === "top") next.y = -epsilon;
  if (edge === "bottom") next.y = 1 + epsilon;
  if (edge === "left") next.x = -epsilon;
  if (edge === "right") next.x = 1 + epsilon;
  return {
    x: clampNumber(next.x, -2, 3, defaultCalloutTailX(type)),
    y: clampNumber(next.y, -2, 3, defaultCalloutTailY(type)),
  };
}



function calloutUnifiedPathD(type, geometry) {
  const outline = calloutClosedOutlinePoints(type, geometry.body);
  const [targetA, targetB] = calloutTailBoundaryTargets(geometry, type);
  if (!outline.length || !targetA || !targetB) return "";
  const hitA = calloutBoundaryHitOnOutline(geometry.tip, targetA, outline) || calloutNearestOutlineHit(targetA, outline);
  const hitB = calloutBoundaryHitOnOutline(geometry.tip, targetB, outline) || calloutNearestOutlineHit(targetB, outline);
  if (!hitA || !hitB) return outlinePointsToPathD(outline);
  const route = calloutOutlineRouteBetween(outline, hitB, hitA);
  return [
    `M${hitA.point.x} ${hitA.point.y}`,
    `L${geometry.tip.x} ${geometry.tip.y}`,
    `L${hitB.point.x} ${hitB.point.y}`,
    route.slice(1).map((point) => `L${point.x} ${point.y}`).join(" "),
    "Z",
  ].filter(Boolean).join(" ");
}



function calloutClosedOutlinePoints(type, body) {
  if (type === "calloutRoundRect") return roundedRectOutlinePoints(body, Math.min(8, Math.max(3, (body.right - body.left) * 0.08)), 8);
  if (type === "calloutOval") return ellipseOutlinePoints(body, 80);
  if (type === "calloutCloud") return cloudOutlinePoints(body, 18);
  return [
    { x: body.left, y: body.top },
    { x: body.right, y: body.top },
    { x: body.right, y: body.bottom },
    { x: body.left, y: body.bottom },
  ].map(roundPoint);
}



function calloutBodySelectionBounds(type) {
  if (type !== "calloutCloud") return { left: 0, top: 0, right: 100, bottom: 60 };
  return boundsFromSvgPoints(cloudOutlinePoints({ left: 0, top: 0, right: 100, bottom: 60 }, 24)) || { left: 0, top: 0, right: 100, bottom: 60 };
}



function boundsFromSvgPoints(points) {
  const valid = (points || []).filter((point) => Number.isFinite(Number(point.x)) && Number.isFinite(Number(point.y)));
  if (!valid.length) return null;
  return {
    left: Math.min(...valid.map((point) => Number(point.x))),
    top: Math.min(...valid.map((point) => Number(point.y))),
    right: Math.max(...valid.map((point) => Number(point.x))),
    bottom: Math.max(...valid.map((point) => Number(point.y))),
  };
}



function roundedRectOutlinePoints(body, radius, stepsPerCorner = 8) {
  const r = Math.max(0, Math.min(radius, (body.right - body.left) / 2, (body.bottom - body.top) / 2));
  const addArc = (points, cx, cy, start, end) => {
    for (let index = 0; index <= stepsPerCorner; index += 1) {
      const angle = start + ((end - start) * index) / stepsPerCorner;
      points.push(roundPoint({ x: cx + Math.cos(angle) * r, y: cy + Math.sin(angle) * r }));
    }
  };
  const points = [];
  addArc(points, body.right - r, body.top + r, -Math.PI / 2, 0);
  addArc(points, body.right - r, body.bottom - r, 0, Math.PI / 2);
  addArc(points, body.left + r, body.bottom - r, Math.PI / 2, Math.PI);
  addArc(points, body.left + r, body.top + r, Math.PI, Math.PI * 1.5);
  return removeAdjacentDuplicatePoints(points);
}



function ellipseOutlinePoints(body, steps = 80) {
  const cx = (body.left + body.right) / 2;
  const cy = (body.top + body.bottom) / 2;
  const rx = Math.max(1, (body.right - body.left) / 2);
  const ry = Math.max(1, (body.bottom - body.top) / 2);
  const points = [];
  for (let index = 0; index < steps; index += 1) {
    const angle = (Math.PI * 2 * index) / steps;
    points.push(roundPoint({ x: cx + Math.cos(angle) * rx, y: cy + Math.sin(angle) * ry }));
  }
  return points;
}



function calloutBoundaryHitOnOutline(tip, target, outline) {
  const rayEnd = calloutBoundaryRayEnd(tip, target);
  let best = null;
  for (let index = 0; index < outline.length; index += 1) {
    const current = outline[index];
    const next = outline[(index + 1) % outline.length];
    const intersection = segmentIntersectionPoint(tip, rayEnd, current, next);
    if (!intersection) continue;
    if (!best || intersection.t < best.t) {
      best = {
        t: intersection.t,
        u: intersection.u,
        index,
        point: roundPoint(intersection),
      };
    }
  }
  return best;
}



function calloutNearestOutlineHit(target, outline) {
  if (!outline.length) return null;
  let best = null;
  for (let index = 0; index < outline.length; index += 1) {
    const current = outline[index];
    const next = outline[(index + 1) % outline.length];
    const dx = Number(next.x) - Number(current.x);
    const dy = Number(next.y) - Number(current.y);
    const lengthSq = dx * dx + dy * dy;
    const u = lengthSq
      ? clampNumber(((Number(target.x) - Number(current.x)) * dx + (Number(target.y) - Number(current.y)) * dy) / lengthSq, 0, 1, 0)
      : 0;
    const point = roundPoint({
      x: Number(current.x) + dx * u,
      y: Number(current.y) + dy * u,
    });
    const distance = squaredDistance(point, target);
    if (!best || distance < best.distance) best = { distance, u, index, point };
  }
  return best
    ? {
        u: best.u,
        index: best.index,
        point: best.point,
      }
    : null;
}



function calloutCloudBoundaryHit(tip, base, body, outline = cloudOutlinePoints(body, 18)) {
  if (outline.length < 2) return null;
  const rayEnd = calloutBoundaryRayEnd(tip, base);
  let best = null;
  for (let index = 0; index < outline.length; index += 1) {
    const current = outline[index];
    const next = outline[(index + 1) % outline.length];
    const intersection = segmentIntersectionPoint(tip, rayEnd, current, next);
    if (!intersection) continue;
    if (!best || intersection.t < best.t) {
      best = {
        t: intersection.t,
        u: intersection.u,
        index,
        point: { x: roundSvgNumber(intersection.x), y: roundSvgNumber(intersection.y) },
      };
    }
  }
  return best;
}



function calloutBoundaryRayEnd(tip, base) {
  const dx = Number(base.x) - Number(tip.x);
  const dy = Number(base.y) - Number(tip.y);
  if (!Number.isFinite(dx) || !Number.isFinite(dy) || (Math.abs(dx) < 0.001 && Math.abs(dy) < 0.001)) return base;
  return {
    x: tip.x + dx * 20,
    y: tip.y + dy * 20,
  };
}



function calloutOutlineRouteBetween(points, startHit, endHit) {
  if (!points.length) return [startHit.point, endHit.point];
  const forward = closedPolylineRoute(points, startHit, endHit, 1);
  const backward = closedPolylineRoute(points, startHit, endHit, -1);
  return polylineLength(forward) >= polylineLength(backward) ? forward : backward;
}



function closedPolylineRoute(points, startHit, endHit, direction = 1) {
  const count = points.length;
  const route = [startHit.point];
  if (!count) {
    route.push(endHit.point);
    return route;
  }
  if (direction > 0) {
    if (startHit.index === endHit.index && endHit.u > startHit.u) {
      route.push(endHit.point);
      return route;
    }
    let vertex = (startHit.index + 1) % count;
    for (let guard = 0; guard <= count + 1; guard += 1) {
      route.push(points[vertex]);
      if (vertex === endHit.index) {
        route.push(endHit.point);
        break;
      }
      vertex = (vertex + 1) % count;
    }
    return route;
  }
  if (startHit.index === endHit.index && endHit.u < startHit.u) {
    route.push(endHit.point);
    return route;
  }
  let vertex = startHit.index;
  const stopVertex = (endHit.index + 1) % count;
  for (let guard = 0; guard <= count + 1; guard += 1) {
    route.push(points[vertex]);
    if (vertex === stopVertex) {
      route.push(endHit.point);
      break;
    }
    vertex = (vertex - 1 + count) % count;
  }
  return route;
}



function polylineLength(points) {
  let length = 0;
  for (let index = 1; index < points.length; index += 1) {
    length += Math.sqrt(squaredDistance(points[index - 1], points[index]));
  }
  return length;
}

export {
  addWorksheetImageDescriptorsToSheet,
  applyHeaderFooterDialog,
  boundsFromSvgPoints,
  calloutBodySelectionBounds,
  calloutBoundaryHitOnOutline,
  calloutBoundaryRayEnd,
  calloutClosedOutlinePoints,
  calloutCloudBoundaryHit,
  calloutCloudBoundaryPoint,
  calloutEffectiveEdge,
  calloutEllipseBoundaryPoint,
  calloutFillOnlyAttrs,
  calloutGeometry,
  calloutNearestOutlineHit,
  calloutOutlineRouteBetween,
  calloutRenderableTailPoint,
  calloutShapePath,
  calloutStrokeOnlyAttrs,
  calloutTailBasePoints,
  calloutTailBoundaryTargets,
  calloutTailFillPathD,
  calloutTailMarkup,
  calloutTailOutlineBasePoints,
  calloutTailOutlinePathD,
  calloutUnifiedPathD,
  calloutViewBoxAttr,
  calloutViewBoxFromPoints,
  calloutWedgeOfficeGeometry,
  clamp01,
  clampNumber,
  closedPolylineRoute,
  cssColorForSvg,
  defaultCalloutTailX,
  defaultCalloutTailY,
  defaultShapeDescriptor,
  defaultStarInnerRatio,
  elbowConnectorRoutePoints,
  elbowConnectorRoutePointsPx,
  ellipseOutlinePoints,
  endingCellForPixelSize,
  endingIndexForPixelSpan,
  explosionBurstPoints,
  explosionSealPoints,
  fileExtension,
  fileExtensionFromImageMimeType,
  fileToDataUrl,
  firstNonZeroLineSegment,
  fitObjectSize,
  handleInsertIconSearch,
  headerFooterDialogHtml,
  headerFooterPathValue,
  headerFooterPlainText,
  headerFooterSectionHasText,
  headerFooterSettingsFromSimpleText,
  imageDimensionsFromSource,
  insert3dModelPlaceholder,
  insertEmbeddedObjectPlaceholder,
  insertEquationObject,
  insertIconDefinitions,
  insertIconDialogHtml,
  insertIconObject,
  insertOfficeAddinObject,
  insertPresetTextObject,
  insertScreenshotObject,
  insertSlicerObject,
  insertSmartArtObject,
  insertTimelineObject,
  insertWorksheetImageDescriptors,
  insertWorksheetImageDescriptorsToSheetGroup,
  insertionPointForObjects,
  installedOfficeAddins,
  isCalloutShapeType,
  isLineShapeObject,
  isLineShapeType,
  linePointToPixelPoint,
  mathShapeText,
  normalizeLineShapeConnections,
  normalizeRotationDegrees,
  normalizeShapeAdjustments,
  normalizeShapeArrowSize,
  normalizeShapeArrowType,
  normalizeShapeCustomPath,
  normalizeShapeFillPattern,
  normalizeShapeGlow,
  normalizeShapeGradient,
  normalizeShapeLineCap,
  normalizeShapeLineCompound,
  normalizeShapeLineDash,
  normalizeShapeLineJoin,
  normalizeShapeLinePoints,
  normalizeShapeOutlineWidth,
  normalizeShapeReflection,
  normalizeShapeShadow,
  normalizeShapeSoftEdges,
  normalizeShapeTextMargin,
  normalizeShapeTextOutlineWidth,
  normalizeShapeTextRuns,
  normalizedHeaderFooterSettings,
  objectDisplaySize,
  officeAddinDefinitions,
  officeAddinsDialogHtml,
  officeDrawingPointsString,
  openHeaderFooterDialog,
  openInsertIconDialog,
  openOfficeAddinsDialog,
  parseSvgViewBoxAttr,
  polylineLength,
  regularPolygonPoints,
  roundRectCornerScale,
  roundSvgTransformNumber,
  roundedRectOutlinePoints,
  setCurrentRibbonShape,
  shapeArrowBaseLength,
  shapeArrowGeometry,
  shapeArrowHalfWidth,
  shapeArrowHeadSvg,
  shapeArrowLineEndpoint,
  shapeArrowMarkerSvg,
  shapeBezierSvgWithArrows,
  shapeBodyFittedToViewBox,
  shapeCanEditPoints,
  shapeCompoundGapColor,
  shapeCompoundLineSvgBody,
  shapeCustomPathSvgD,
  shapeDescriptorForType,
  shapeEffectRadius,
  shapeEffectX,
  shapeEffectY,
  shapeEffectiveOutlineWidth,
  shapeEffectsSvg,
  shapeFreeformSvgWithArrows,
  shapeFreeformSvgWithArrowsPx,
  shapeGradientDefSvg,
  shapeGradientVector,
  shapeGroupForType,
  shapeLineArrowHeadsSvg,
  shapeLineDashArray,
  shapeLineSvgMarkup,
  shapeLineSvgWithArrows,
  shapePatternDefSvg,
  shapePlainTextFromRuns,
  shapePolylineSvgWithArrows,
  shapePresetForType,
  shapeReflectionSvgMarkup,
  shapeShadowRenderModel,
  shapeSvgMarkup,
  shapeSvgStrokeLayer,
  shapeTextRunStyleKey,
  shapeTextSvgMarkup,
  shapeVisibleLinePoints,
  star4Points,
  starPointCountForType,
  starPoints,
  svgColorParts,
  worksheetIconSvg,
  worksheetSvgObjectDescriptor,
};
