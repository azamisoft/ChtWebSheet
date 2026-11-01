// Feature module; shared bindings are imported explicitly.
import {
  $appAlert,
  $statusText,
  $warnings,
  APP_COMPANY_NAME,
  APP_DISPLAY_NAME,
  APP_SHORT_NAME,
  APP_VERSION,
  AUTO_STANDALONE_RUNTIME_BASE_URL,
  DEFAULT_STANDALONE_REMOTE_RUNTIME_BASE_URL,
  DEFAULT_STANDALONE_RUNTIME_BASE_URL,
  DIALOG_CHROME_ROOT_SELECTOR,
  DIALOG_CHROME_TITLE_SELECTOR,
  SERVER_DOWNLOAD_PAGE_PATH,
  SERVER_VERSION_INFO_PATH,
  STANDALONE_REMOTE_RUNTIME_BASE_URL_STORAGE_KEY,
  STANDALONE_RUNTIME_BASE_URL_STORAGE_KEY,
  STANDALONE_RUNTIME_CHANNEL,
  STANDALONE_RUNTIME_VERSION,
  createIcons,
  icons,
  state,
} from "../../app-core.js";
import {
  closeFindDialog,
} from "../data/query-editor.js";
import {
  accessKeyFromKeyboardEvent,
  closeDataDialog,
} from "../drawing/charts-and-insertion.js";
import {
  closeSymbolDialog,
} from "../editing/ink-and-clipboard.js";
import {
  isElementImeCompositionEvent,
} from "../editing/keyboard-and-paste.js";
import {
  closeCellContextMenu,
  closeCellStructureDialog,
  closeFormatCellsDialog,
  closeHeaderSizeDialog,
  closeLinkDialog,
  closeNameDialog,
} from "./annotations-and-cell-format.js";
import {
  closeGoToDialog,
  closeHomeSeriesDialog,
  closeHomeWorksheetFillDialog,
  closeSpecialDialog,
} from "./printing-and-home-tools.js";
import {
  closeFileSideMenu,
  closeObjectLinkDialog,
} from "./ribbon-and-links.js";
import {
  closeSheetTabContextMenu,
  closeZoomDialog,
} from "./sheet-protection-and-translation.js";
import {
  appHostedAssetUrl,
  defaultStandaloneRuntimeBaseUrl,
  downloadAssetUrlFromRuntimeBase,
  isRemoteRuntimeBaseUrl,
  normalizeRuntimeBaseUrl,
  runtimeVersionedBaseUrlFromCurrent,
  uniqueRuntimeBaseUrls,
} from "../workbook/export-model-and-measurement.js";
import {
  isElapsedTimeFormat,
  primaryExcelFormatSection,
} from "../workbook/saving-and-export.js";
import {
  menuAccessKeyFromLabel,
} from "../workbook/views-and-macros.js";


function standaloneCss() {
  return `
@font-face{font-family:"MS Pゴシック";src:local("MS PGothic"),local("ＭＳ Ｐゴシック"),local("MS Pゴシック")}@font-face{font-family:"MS P明朝";src:local("MS PMincho"),local("ＭＳ Ｐ明朝"),local("MS P明朝")}
html,body{height:100%;overflow:hidden}body{margin:0;background:#eef2f7;color:#1d2433;font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
.websheet{height:100vh;min-height:0;overflow:hidden;display:flex;flex-direction:column}
.wx-topbar{min-height:56px;padding:10px 16px;background:#fff;border-bottom:1px solid #d9e0ea;display:flex;align-items:center;justify-content:space-between;gap:12px}
.wx-topbar span{display:block;margin-top:2px;color:#617085;font-size:13px}.wx-topbar button,.wx-formula button{border:1px solid #b7c1d1;background:#fff;border-radius:8px;min-height:32px;padding:0 12px;cursor:pointer}
.wx-formula{padding:8px 16px;background:#fff;border-bottom:1px solid #d9e0ea;display:grid;grid-template-columns:64px minmax(0,1fr)72px;gap:8px}
#wx-cell-label{display:inline-flex;align-items:center;justify-content:center;border:1px solid #d9e0ea;background:#f8fafc;border-radius:8px;font-weight:700}
#wx-formula-input{border:1px solid #b7c1d1;border-radius:8px;padding:0 10px}
#wx-sheet{flex:1;min-height:0;overflow-x:hidden;overflow-y:auto;padding:0}
.wx-footer{--wx-tabs-width:60%;flex:0 0 auto;min-height:32px;border-top:1px solid #d9e0ea;background:#f7f9fc;display:grid;grid-template-columns:minmax(150px,var(--wx-tabs-width)) 8px minmax(220px,1fr)}
.wx-tabs-pane{display:flex;align-items:flex-end;min-width:0;border-right:1px solid #d9e0ea}#wx-tabs{display:flex;align-items:center;height:31px;min-width:0;max-width:100%;overflow-x:auto;scrollbar-width:none}#wx-tabs::-webkit-scrollbar{display:none}
#wx-tabs button{border:0;border-right:1px solid #d9e0ea;background:#f7f9fc;color:#344054;min-height:28px;padding:0 14px;cursor:pointer;white-space:nowrap;flex:0 0 auto}#wx-tabs button.active{background:#fff;color:#111827;font-weight:700;box-shadow:inset 0 -2px 0 #0f766e}
#wx-add{width:34px;min-width:34px;height:31px;border:0;border-right:1px solid #d9e0ea;background:#f7f9fc;color:#475467;font-size:18px;cursor:pointer}#wx-splitter{cursor:col-resize;background:#e5eaf2;border-inline:1px solid #d0d7e2}#wx-footer-right{min-width:0;display:grid;grid-template-rows:16px 16px;background:#f7f9fc}#wx-hscroll{overflow-x:auto;overflow-y:hidden;height:16px;border-bottom:1px solid #d9e0ea}#wx-hspacer{height:1px}
	#wx-view-toolbar{min-height:16px;display:flex;align-items:center;justify-content:flex-end;gap:8px;padding:0 8px 0 6px}#wx-view-buttons{display:inline-flex;align-items:center;gap:1px}#wx-view-buttons button{width:18px;height:14px;min-height:14px;border:1px solid transparent;background:transparent;color:#475467;border-radius:3px;padding:0;white-space:nowrap;cursor:pointer;display:inline-flex;align-items:center;justify-content:center}#wx-view-buttons button:hover,#wx-view-buttons button.active{background:#fff;border-color:#cbd5e1;color:#1d4ed8}.wx-view-mark{width:10px;height:10px;display:block;position:relative;color:currentColor}.wx-view-mark-normal{border:1px solid currentColor;background:linear-gradient(currentColor,currentColor) 50% 0/1px 100% no-repeat,linear-gradient(currentColor,currentColor) 0 50%/100% 1px no-repeat}.wx-view-mark-page{border:1px solid currentColor;background:linear-gradient(currentColor,currentColor) 0 2px/100% 1px no-repeat,linear-gradient(currentColor,currentColor) 2px 5px/6px 1px no-repeat}.wx-view-mark-break{border:1px dashed currentColor}.wx-view-mark-break:before,.wx-view-mark-break:after{content:"";position:absolute;background:currentColor}.wx-view-mark-break:before{left:4px;top:1px;width:1px;height:6px}.wx-view-mark-break:after{left:1px;top:4px;width:6px;height:1px}#wx-zoom{display:inline-flex;align-items:center;gap:4px;min-width:156px}#wx-zoom button{width:14px;height:14px;border:1px solid #cbd5e1;background:#fff;color:#475467;border-radius:50%;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;font-size:11px;padding:0}#wx-zoom-slider{width:80px;height:14px;accent-color:#0f766e}#wx-zoom-label{min-width:36px;color:#475467;font-size:11px;text-align:right;border:1px solid transparent;border-radius:3px;padding:0 2px;cursor:pointer;line-height:14px}#wx-zoom-label:hover,#wx-zoom-label:focus-visible{background:#fff;border-color:#cbd5e1;color:#1d4ed8;outline:none}#wx-zoom-dialog{position:fixed;z-index:1100;width:264px;padding:10px 12px 12px;border:1px solid #9aa4b2;background:#fff;box-shadow:0 12px 30px rgba(15,23,42,.2);color:#1f2937;font-size:13px}.wx-zoom-dialog-title{margin:0 0 10px;font-size:14px;font-weight:700}.wx-zoom-dialog-section-title{margin-bottom:6px;color:#344054;font-weight:600}.wx-zoom-dialog-options{display:grid;gap:5px}.wx-zoom-dialog-option{display:flex;align-items:center;gap:6px;min-height:22px;white-space:nowrap}.wx-zoom-dialog-option input[type=radio]{margin:0}.wx-zoom-dialog-input{width:64px;height:24px;border:1px solid #aeb7c6;padding:0 6px;text-align:right}.wx-zoom-dialog-input:focus{outline:2px solid #bfdbfe;outline-offset:-1px}.wx-zoom-dialog-actions{display:flex;justify-content:flex-end;gap:8px;margin-top:12px}.wx-zoom-dialog-button{min-width:76px;min-height:28px;border:1px solid #aeb7c6;background:#fff;color:#1f2937;cursor:pointer}.wx-zoom-dialog-button:hover,.wx-zoom-dialog-button:focus-visible{border-color:#2563eb;outline:2px solid #bfdbfe;outline-offset:-1px}.wx-zoom-dialog-button.primary{border-color:#2563eb}.menu-access-key{text-decoration:underline;text-underline-offset:2px}.menu-access-key-token{display:inline-block;white-space:nowrap}
.sheet-grid{display:grid;position:relative;width:max-content;min-width:100%;background:#fff;box-shadow:0 1px 3px rgba(15,23,42,.08)}
.sheet-grid.view-page-layout{margin:18px auto 28px;min-width:auto;box-shadow:0 0 0 1px #cfd8e3,0 16px 34px rgba(15,23,42,.14)}.sheet-grid.view-page-layout.hide-white-space{box-shadow:0 0 0 1px #d9e0ea}
.sheet-grid.view-page-break{background:#fff;box-shadow:inset 0 0 0 2px #8bb8ff,0 1px 3px rgba(15,23,42,.08)}
	.sheet-page-break-page,.sheet-page-break-outside-print-area{position:absolute;pointer-events:none;box-sizing:border-box}.sheet-page-break-page{z-index:2;border:0;background:transparent;overflow:hidden}.sheet-page-break-outside-print-area{z-index:3;background:repeating-linear-gradient(45deg,rgba(37,99,235,.08) 0,rgba(37,99,235,.08) 8px,rgba(191,219,254,.16) 8px,rgba(191,219,254,.16) 16px)}.sheet-page-break-outside-label{position:absolute;left:14px;top:12px;color:rgba(30,64,175,.42);font:700 16px/1.2 "Yu Gothic UI",Meiryo,sans-serif;user-select:none}.sheet-page-break-label{position:absolute;z-index:4;pointer-events:none;color:rgba(37,99,235,.28);font:700 42px/1 "Yu Gothic UI",Meiryo,sans-serif;user-select:none}
.sheet-canvas{position:absolute;top:0;left:0;z-index:0;pointer-events:none}.sheet-ink-draft,.sheet-ink-layer{position:absolute;z-index:90000;overflow:visible;pointer-events:none}.sheet-ink-stroke{vector-effect:non-scaling-stroke}
.grid-header{border:1px solid #d9e0ea}.grid-header,.sheet-cell{padding:0 3px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;vertical-align:middle;font-size:13px;line-height:1.2;letter-spacing:0;position:relative}
.grid-header{background:#f3f6fa;color:#526173;text-align:center;font-weight:600;user-select:none;text-overflow:clip}.column-header{position:sticky;top:0;z-index:100000}.row-header{position:sticky;left:0;z-index:100000}.corner-header{position:sticky;top:0;left:0;z-index:100001}.header-label{display:block;min-width:0;max-width:100%;overflow:hidden;text-overflow:clip;white-space:nowrap;pointer-events:none}.sheet-cell{z-index:1;border-right:1px solid transparent;border-bottom:1px solid transparent;background:transparent;background-clip:padding-box;display:flex;align-items:center;user-select:none}.sheet-cell.has-fill{background-clip:border-box}.sheet-cell.has-fill,.sheet-cell.hide-grid-right{border-right-color:transparent}.sheet-cell.has-fill,.sheet-cell.hide-grid-bottom{border-bottom-color:transparent}.sheet-grid.hide-gridlines .sheet-cell{border-right-color:transparent;border-bottom-color:transparent}.corner-header,.row-header{width:100%;min-width:0;max-width:100%;padding:0}
.sheet-cell.text-overflow{overflow:visible;z-index:2}.sheet-cell.text-overflow .cell-content{pointer-events:none}.cell-content{display:block;width:100%;flex:0 0 auto;overflow:hidden;text-overflow:clip;white-space:inherit}.sheet-cell:not([contenteditable=true]) .cell-content{opacity:0}
.sheet-image,.sheet-image-link{position:absolute;display:block;box-sizing:border-box;z-index:3}.sheet-image{object-fit:fill;pointer-events:auto;user-select:none}.sheet-image.sheet-shape:not(.is-line-shape){overflow:visible}.sheet-shape-render{position:absolute;inset:0;display:block;width:100%;height:100%;pointer-events:none;user-select:none}.sheet-shape-text-layer{position:absolute;inset:0;box-sizing:border-box;display:flex;flex-direction:column;min-width:0;min-height:0;overflow:hidden;line-height:1.22;white-space:pre-wrap;overflow-wrap:break-word;word-break:normal;pointer-events:none;user-select:none}.sheet-shape-text-block{width:100%;min-width:0}.sheet-picture-crop{overflow:hidden;object-fit:initial}.sheet-picture-crop-content{position:absolute;display:block;max-width:none;max-height:none;pointer-events:none;user-select:none}.sheet-image-link{pointer-events:auto}.sheet-image-link .sheet-image{position:static;width:100%;height:100%}.sheet-image-link .sheet-picture-crop{position:relative}.sheet-image.selected{outline:2px solid #22c55e;outline-offset:1px}.wx-selection-frame{position:absolute;z-index:5;pointer-events:none;box-sizing:border-box;border:2px solid #22c55e;background:rgba(34,197,94,.08)}.wx-selection-frame.is-active::after{content:"";position:absolute;right:-4px;bottom:-4px;width:7px;height:7px;background:#22c55e;border:1px solid #fff;box-sizing:border-box}
.sheet-cell[contenteditable=true]{background:transparent;box-sizing:border-box;cursor:text;display:block;height:var(--cell-edit-overlay-height,100%);isolation:isolate;min-height:100%;overflow:visible;padding-top:var(--cell-edit-padding-top,2px);text-overflow:clip;white-space:pre;user-select:text;scrollbar-width:none;z-index:40}.sheet-cell[contenteditable=true]::before{content:"";position:absolute;z-index:-1;left:var(--cell-edit-overlay-left,0);top:0;width:var(--cell-edit-overlay-width,100%);height:var(--cell-edit-overlay-height,100%);background:var(--cell-fill-color,#fff);pointer-events:none}.sheet-cell[contenteditable=true]::-webkit-scrollbar{width:0;height:0;display:none}.sheet-cell[contenteditable=true] .cell-content{opacity:1;overflow:visible;pointer-events:auto}.sheet-cell[contenteditable=true]:focus{outline:2px solid #22c55e;outline-offset:-2px;z-index:41}.sheet-cell.selected{outline:2px solid #22c55e;outline-offset:-2px;z-index:4}
.formula-cell:after{content:"";position:absolute;right:3px;top:3px;width:0;height:0;border-left:6px solid transparent;border-top:6px solid #2563eb}
.cell-annotation-marker{position:absolute;top:0;right:0;z-index:6;width:0;height:0;pointer-events:none;border-left:8px solid transparent}.cell-annotation-marker.is-comment{border-top:8px solid #8064a2}.cell-annotation-marker.is-note{border-top:8px solid #c00000}.sheet-cell.has-comment.has-note .cell-annotation-marker.is-comment{right:9px}
a{color:#1d4ed8}.has-validation-dropdown .cell-content{padding-right:18px}.cell-validation-dropdown-button{position:absolute;right:0;top:50%;z-index:3;display:none;width:16px;height:18px;transform:translateY(-50%);border:1px solid #aeb7c6;border-radius:0;background:linear-gradient(#fff,#e8eef7);color:#1f2937;cursor:default;padding:0}.sheet-cell.selected .cell-validation-dropdown-button,.cell-validation-dropdown-button:focus-visible{display:block}.cell-validation-dropdown-button:after{content:"";position:absolute;left:50%;top:50%;width:0;height:0;transform:translate(-50%,-25%);border-left:4px solid transparent;border-right:4px solid transparent;border-top:5px solid currentColor}.cell-validation-dropdown-button:hover,.cell-validation-dropdown-button:focus-visible{border-color:#22c55e;outline:none;background:#eef4ff}.validation-dropdown-popup{position:fixed;z-index:2140;max-height:220px;overflow-y:auto;border:1px solid #7f9db9;background:#fff;box-shadow:0 6px 18px rgba(15,23,42,.18);color:#111827;font:12px "Yu Gothic UI","Meiryo UI","Segoe UI",sans-serif}.validation-dropdown-choice{display:block;width:100%;min-height:22px;padding:2px 8px;border:0;background:#fff;color:inherit;text-align:left;white-space:nowrap}.validation-dropdown-choice:hover,.validation-dropdown-choice:focus,.validation-dropdown-choice.is-selected{background:#dcfce7;outline:none}
`;
}



function initDialogChromeObserver() {
  decorateDialogChrome(document);
  if (state.dialogChromeObserver || typeof MutationObserver === "undefined") return;
  state.dialogChromeObserver = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) decorateDialogChrome(node);
      });
    });
  });
  state.dialogChromeObserver.observe(document.body, { childList: true, subtree: true });
}



function decorateDialogChrome(root) {
  const dialogs = [];
  if (root?.matches?.(DIALOG_CHROME_ROOT_SELECTOR)) dialogs.push(root);
  root?.querySelectorAll?.(DIALOG_CHROME_ROOT_SELECTOR).forEach((dialog) => dialogs.push(dialog));
  dialogs.forEach(decorateSingleDialogChrome);
}



function decorateSingleDialogChrome(dialog) {
  if (!dialog || dialog.dataset.dialogChromeEnhanced === "true") return;
  const titlebar = dialog.querySelector(DIALOG_CHROME_TITLE_SELECTOR);
  if (!titlebar) return;
  titlebar.dataset.dialogDragHandle = "true";
  titlebar.setAttribute("title", "ドラッグしてダイアログを移動");
  if (!titlebar.querySelector("[data-dialog-chrome-close]")) {
    const closeButton = document.createElement("button");
    closeButton.type = "button";
    closeButton.className = "dialog-chrome-close";
    closeButton.dataset.dialogChromeClose = "true";
    closeButton.setAttribute("aria-label", "閉じる");
    closeButton.textContent = "×";
    titlebar.append(closeButton);
  }
  dialog.dataset.dialogChromeEnhanced = "true";
}



function handleDialogChromeCloseClick(event) {
  event.preventDefault();
  event.stopPropagation();
  const dialog = event.currentTarget.closest(DIALOG_CHROME_ROOT_SELECTOR);
  closeDialogFromChrome(dialog);
}



function closeDialogFromChrome(dialog) {
  if (!dialog) return;
  if (dialog.closest("#messageBoxOverlay")) {
    closeMessageBox(state.messageBoxHandlers?.cancelResult || "ok", true);
    return;
  }
  if (dialog.id === "zoomDialog") {
    closeZoomDialog();
    return;
  }
  const overlay = dialog.closest(".excel-dialog-overlay, .format-dialog-overlay, .cell-structure-dialog-overlay") || dialog.parentElement;
  const overlayId = overlay?.id || "";
  const dialogId = dialog.id || "";
  switch (overlayId || dialogId) {
    case "dataDialog":
      closeDataDialog();
      return;
    case "formatCellsDialogOverlay":
    case "formatCellsDialog":
      closeFormatCellsDialog();
      return;
    case "cellStructureDialogOverlay":
    case "cellStructureDialog":
      closeCellStructureDialog();
      return;
    case "headerSizeDialogOverlay":
    case "headerSizeDialog":
      closeHeaderSizeDialog();
      return;
    case "nameDialogOverlay":
    case "nameDialog":
      closeNameDialog();
      return;
    case "linkDialogOverlay":
    case "linkDialog":
      closeLinkDialog();
      return;
    case "objectLinkDialogOverlay":
    case "objectLinkDialog":
      closeObjectLinkDialog();
      return;
    case "homeSeriesDialogOverlay":
    case "homeSeriesDialog":
      closeHomeSeriesDialog();
      return;
    case "homeWorksheetFillDialogOverlay":
    case "homeWorksheetFillDialog":
      closeHomeWorksheetFillDialog();
      return;
    case "goToDialogOverlay":
    case "goToDialog":
      closeGoToDialog();
      return;
    case "specialDialogOverlay":
    case "specialDialog":
      closeSpecialDialog();
      return;
    case "findDialogOverlay":
    case "findDialog":
      closeFindDialog();
      return;
    case "symbolDialogOverlay":
    case "symbolDialog":
      closeSymbolDialog();
      return;
    case "helpDialogOverlay":
      closeHelpDialog();
      return;
    case "versionDialogOverlay":
      closeVersionDialog();
      return;
    default:
      overlay?.remove?.();
      dialog.remove?.();
  }
}



function startDialogDrag(event) {
  if (event.button !== 0) return;
  if (event.target.closest?.("[data-dialog-chrome-close], button, input, select, textarea, a, [contenteditable='true']")) return;
  const titlebar = event.currentTarget;
  const dialog = titlebar.closest(DIALOG_CHROME_ROOT_SELECTOR);
  if (!dialog) return;
  const rect = dialog.getBoundingClientRect();
  if (!rect.width || !rect.height) return;
  event.preventDefault();
  event.stopPropagation();
  dialog.style.position = "fixed";
  dialog.style.left = `${Math.round(rect.left)}px`;
  dialog.style.top = `${Math.round(rect.top)}px`;
  dialog.style.margin = "0";
  dialog.style.transform = "none";
  dialog.style.width = `${Math.round(rect.width)}px`;
  state.dialogDrag = {
    dialog,
    startX: event.clientX,
    startY: event.clientY,
    left: rect.left,
    top: rect.top,
    width: rect.width,
    height: rect.height,
  };
  document.body.classList.add("is-dialog-dragging");
  titlebar.setPointerCapture?.(event.pointerId);
}



function updateDialogDrag(event) {
  const drag = state.dialogDrag;
  if (!drag?.dialog?.isConnected) return;
  const nextLeft = drag.left + event.clientX - drag.startX;
  const nextTop = drag.top + event.clientY - drag.startY;
  drag.dialog.style.left = `${Math.round(nextLeft)}px`;
  drag.dialog.style.top = `${Math.round(nextTop)}px`;
}



function finishDialogDrag() {
  if (!state.dialogDrag) return;
  state.dialogDrag = null;
  document.body.classList.remove("is-dialog-dragging");
}



function setStatus(message) {
  $statusText.text(message);
}



function serverVersionInfoUrl() {
  if (typeof window === "undefined") return SERVER_VERSION_INFO_PATH;
  try {
    if (/^https?:$/i.test(window.location.protocol)) {
      return appHostedAssetUrl(SERVER_VERSION_INFO_PATH);
    }
  } catch {
    // Fall through to the standalone runtime server if the current origin is not usable.
  }
  const remoteBase = window.__WEBSHEET_STANDALONE_CONFIG__?.remoteBaseUrl ||
    (() => {
      try {
        return localStorage.getItem(STANDALONE_REMOTE_RUNTIME_BASE_URL_STORAGE_KEY);
      } catch {
        return "";
      }
    })();
  if (isRemoteRuntimeBaseUrl(remoteBase)) {
    try {
      return downloadAssetUrlFromRuntimeBase(SERVER_VERSION_INFO_PATH, remoteBase);
    } catch {
      // Use the default path below.
    }
  }
  const hostedUrl = appHostedAssetUrl(SERVER_VERSION_INFO_PATH);
  return isRemoteRuntimeBaseUrl(hostedUrl)
    ? hostedUrl
    : downloadAssetUrlFromRuntimeBase(SERVER_VERSION_INFO_PATH, DEFAULT_STANDALONE_REMOTE_RUNTIME_BASE_URL);
}



function versionInfoPayload(info = {}) {
  const payload = info && typeof info === "object" ? info : {};
  return payload.versionJson && typeof payload.versionJson === "object" ? payload.versionJson : payload;
}



function versionInfoAppVersion(info = {}) {
  const payload = versionInfoPayload(info);
  return String(payload.appVersion || payload.version || "").trim();
}



function versionInfoRuntimeVersion(info = {}) {
  const payload = versionInfoPayload(info);
  return String(payload.runtimeVersion || payload.runtime?.version || payload.version || "").trim();
}



function versionInfoUpdatedAt(info = {}) {
  const payload = versionInfoPayload(info);
  return String(payload.updatedAt || payload.runtimeUpdatedAt || payload.runtime?.updatedAt || "").trim();
}



function versionInfoRuntimeUpdatedAt(info = {}) {
  const payload = versionInfoPayload(info);
  return String(payload.runtimeUpdatedAt || payload.runtime?.updatedAt || payload.updatedAt || "").trim();
}



function activeRuntimeVersionInfo() {
  const source = runtimeExecutionSource();
  if (source === "local" && window.__WEBSHEET_LOCAL_RUNTIME_INFO__) return window.__WEBSHEET_LOCAL_RUNTIME_INFO__;
  if (source === "remote" && window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__) return window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__;
  if (state.serverVersionInfo) return state.serverVersionInfo;
  if (window.__WEBSHEET_LOCAL_RUNTIME_INFO__) return window.__WEBSHEET_LOCAL_RUNTIME_INFO__;
  return window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__ || {};
}



function currentRuntimeVersion() {
  return versionInfoRuntimeVersion(activeRuntimeVersionInfo()) || STANDALONE_RUNTIME_VERSION;
}



function currentRuntimeUpdatedAt() {
  return versionInfoRuntimeUpdatedAt(activeRuntimeVersionInfo());
}



function requiredRuntimeVersion() {
  return versionInfoRuntimeVersion(state.serverVersionInfo) || currentRuntimeVersion();
}



function requiredRuntimeUpdatedAt() {
  return versionInfoRuntimeUpdatedAt(state.serverVersionInfo) || currentRuntimeUpdatedAt();
}



function normalizeVersionInfo(info = {}) {
  const version = versionInfoAppVersion(info) || APP_VERSION;
  const runtimeVersion = versionInfoRuntimeVersion(info) || version || STANDALONE_RUNTIME_VERSION;
  const updatedAt = versionInfoUpdatedAt(info);
  const runtimeUpdatedAt = versionInfoRuntimeUpdatedAt(info) || updatedAt;
  const payload = versionInfoPayload(info);
  return {
    appName: String(payload.appName || APP_DISPLAY_NAME),
    company: String(payload.company || APP_COMPANY_NAME),
    version,
    runtimeVersion,
    updatedAt,
    runtimeUpdatedAt,
    downloadPage: String(payload.downloadPage || SERVER_DOWNLOAD_PAGE_PATH),
    packages: payload.packages && typeof payload.packages === "object" ? payload.packages : {},
    raw: info,
  };
}



function versionInfoIsMismatch(info = state.serverVersionInfo) {
  if (!info) return false;
  const serverVersion = versionInfoRuntimeVersion(info);
  const activeVersion = currentRuntimeVersion();
  if (!serverVersion || !activeVersion || activeVersion === STANDALONE_RUNTIME_VERSION) return false;
  if (compareRuntimeVersions(serverVersion, activeVersion) !== 0) return true;
  const serverUpdatedAt = runtimeUpdatedAtTimestamp(versionInfoRuntimeUpdatedAt(info));
  const activeUpdatedAt = runtimeUpdatedAtTimestamp(currentRuntimeUpdatedAt());
  return Boolean(serverUpdatedAt && activeUpdatedAt && serverUpdatedAt > activeUpdatedAt);
}



async function refreshServerVersionInfo(options = {}) {
  if (state.serverVersionStatus === "loading" && !options.force) return state.serverVersionInfo;
  state.serverVersionStatus = "loading";
  syncVersionRibbonBadge();
  renderVersionDialogContent();
  try {
    const response = await fetch(`${serverVersionInfoUrl()}?t=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    state.serverVersionInfo = normalizeVersionInfo(await response.json());
    state.serverVersionStatus = "loaded";
    state.versionMismatch = versionInfoIsMismatch(state.serverVersionInfo);
  } catch (error) {
    state.serverVersionStatus = "error";
    state.serverVersionInfo = null;
    state.versionMismatch = false;
    console.warn("Server version info could not be loaded.", error);
  }
  syncVersionRibbonBadge();
  renderVersionDialogContent();
  return state.serverVersionInfo;
}



function runtimeVersionParts(value) {
  return String(value || "").split(/[.-]/).map((part) => {
    const number = Number.parseInt(part, 10);
    return Number.isFinite(number) ? number : 0;
  });
}



function compareRuntimeVersions(left, right) {
  const a = runtimeVersionParts(left);
  const b = runtimeVersionParts(right);
  const length = Math.max(a.length, b.length, 3);
  for (let index = 0; index < length; index += 1) {
    const delta = (a[index] || 0) - (b[index] || 0);
    if (delta) return delta;
  }
  return 0;
}



function runtimeUpdatedAtTimestamp(value) {
  const text = String(value || "").trim();
  if (!text) return 0;
  const time = Date.parse(text);
  return Number.isFinite(time) ? time : 0;
}



function runtimeInfoIsUsable(info) {
  const version = versionInfoRuntimeVersion(info);
  if (!version) return false;
  const versionCompare = compareRuntimeVersions(version, requiredRuntimeVersion());
  if (versionCompare < 0) return false;
  if (versionCompare > 0) return true;
  const requiredUpdatedAt = runtimeUpdatedAtTimestamp(requiredRuntimeUpdatedAt());
  if (!requiredUpdatedAt) return true;
  const actualUpdatedAt = runtimeUpdatedAtTimestamp(versionInfoRuntimeUpdatedAt(info));
  return Boolean(actualUpdatedAt && actualUpdatedAt >= requiredUpdatedAt);
}



function inferredUserRuntimeBaseUrlsFromCurrentFile() {
  const bases = [];
  try {
    const currentUrl = new URL(window.location.href);
    if (currentUrl.protocol !== "file:") return bases;
    const path = currentUrl.pathname || "";
    const version = encodeURIComponent(requiredRuntimeVersion());
    const add = (root) => {
      bases.push(`${root}/${STANDALONE_RUNTIME_CHANNEL}`);
      if (version) bases.push(`${root}/${version}`);
    };
    const macMatch = path.match(/^\/Users\/([^/]+)\//);
    if (macMatch) add(`file:///Users/${macMatch[1]}/Library/Application%20Support/WebSheet/runtime`);
    const windowsMatch = path.match(/^\/([A-Za-z]:)\/Users\/([^/]+)\//);
    if (windowsMatch) add(`file:///${windowsMatch[1]}/Users/${windowsMatch[2]}/AppData/Roaming/WebSheet/runtime`);
    const linuxMatch = path.match(/^\/home\/([^/]+)\//);
    if (linuxMatch) add(`file:///home/${linuxMatch[1]}/.local/share/WebSheet/runtime`);
  } catch {
    // Keep configured candidates when the current document URL is not parseable.
  }
  return bases;
}



function configuredLocalRuntimeBaseUrls() {
  const config = window.__WEBSHEET_STANDALONE_CONFIG__ || {};
  let customRuntimeBaseUrl = "";
  try {
    customRuntimeBaseUrl = localStorage.getItem(STANDALONE_RUNTIME_BASE_URL_STORAGE_KEY) || "";
  } catch {
    // localStorage can be unavailable for some local files.
  }
  const defaults = [];
  const defaultBaseUrl = String(DEFAULT_STANDALONE_RUNTIME_BASE_URL || "").trim();
  if (defaultBaseUrl && defaultBaseUrl !== AUTO_STANDALONE_RUNTIME_BASE_URL) {
    defaults.push(defaultBaseUrl, runtimeVersionedBaseUrlFromCurrent(defaultBaseUrl));
  }
  return uniqueRuntimeBaseUrls([
    ...inferredUserRuntimeBaseUrlsFromCurrentFile(),
    customRuntimeBaseUrl,
    customRuntimeBaseUrl ? runtimeVersionedBaseUrlFromCurrent(customRuntimeBaseUrl) : "",
    config.localBaseUrl,
    config.versionedLocalBaseUrl,
    ...defaults,
    defaultStandaloneRuntimeBaseUrl(),
    runtimeVersionedBaseUrlFromCurrent(defaultStandaloneRuntimeBaseUrl()),
  ]);
}



function canProbeLocalRuntimeVersion() {
  try {
    return window.location.protocol === "file:" || window.__WEBSHEET_STANDALONE__ === true;
  } catch {
    return false;
  }
}



async function fetchLocalRuntimeVersionJson(base) {
  const normalized = normalizeRuntimeBaseUrl(base);
  if (!normalized) return null;
  try {
    const response = await fetch(`${normalized}/version.json?t=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) return null;
    const info = await response.json();
    return info && typeof info === "object" ? info : null;
  } catch {
    return null;
  }
}



function probeLocalRuntimeVersion(base) {
  return new Promise((resolve) => {
    const normalized = normalizeRuntimeBaseUrl(base);
    if (!normalized) {
      resolve(null);
      return;
    }
    let done = false;
    const previousProbe = window.__WEBSHEET_RUNTIME_VERSION_PROBE__;
    const previousInfo = window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__;
    const script = document.createElement("script");
    const timer = window.setTimeout(() => finish(null), 1200);
    function finish(info) {
      if (done) return;
      done = true;
      window.clearTimeout(timer);
      window.__WEBSHEET_RUNTIME_VERSION_PROBE__ = previousProbe;
      script.remove();
      resolve(info || null);
    }
    window.__WEBSHEET_RUNTIME_VERSION_PROBE__ = (info) => finish(info || null);
    script.src = `${normalized}/version.js`;
    script.async = true;
    script.onload = () => {
      const nextInfo = window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__;
      finish(nextInfo && nextInfo !== previousInfo ? nextInfo : null);
    };
    script.onerror = () => finish(null);
    document.head.appendChild(script);
  });
}



async function readLocalRuntimeVersion(base) {
  return await fetchLocalRuntimeVersionJson(base) || await probeLocalRuntimeVersion(base);
}



async function refreshLocalRuntimeInfo(options = {}) {
  const currentStatus = localRuntimeStatus();
  if (!options.force && (currentStatus === "current" || currentStatus === "outdated")) {
    return window.__WEBSHEET_LOCAL_RUNTIME_INFO__ || null;
  }
  if (!canProbeLocalRuntimeVersion()) return window.__WEBSHEET_LOCAL_RUNTIME_INFO__ || null;
  window.__WEBSHEET_LOCAL_RUNTIME_STATUS__ = "checking";
  syncVersionRibbonBadge();
  renderVersionDialogContent();
  const candidates = configuredLocalRuntimeBaseUrls();
  window.__WEBSHEET_LOCAL_RUNTIME_CANDIDATES__ = candidates.slice();
  let firstInfo = null;
  let firstBase = "";
  for (const base of candidates) {
    const info = await readLocalRuntimeVersion(base);
    if (info && !firstInfo) {
      firstInfo = info;
      firstBase = base;
    }
    if (runtimeInfoIsUsable(info)) {
      window.__WEBSHEET_LOCAL_RUNTIME_INFO__ = info;
      window.__WEBSHEET_LOCAL_RUNTIME_BASE_URL__ = base;
      window.__WEBSHEET_LOCAL_RUNTIME_STATUS__ = "current";
      syncVersionRibbonBadge();
      renderVersionDialogContent();
      return info;
    }
  }
  window.__WEBSHEET_LOCAL_RUNTIME_INFO__ = firstInfo || null;
  window.__WEBSHEET_LOCAL_RUNTIME_BASE_URL__ = firstBase;
  window.__WEBSHEET_LOCAL_RUNTIME_STATUS__ = firstInfo ? "outdated" : "missing";
  syncVersionRibbonBadge();
  renderVersionDialogContent();
  return firstInfo;
}



function syncVersionRibbonBadge() {
  const shouldSuggestRuntimeDownload = versionDialogShouldSuggestDownload();
  document.querySelectorAll(".ribbon-version-update-badge").forEach((badge) => {
    badge.classList.toggle("hidden", !state.versionMismatch && !shouldSuggestRuntimeDownload);
  });
  document.querySelectorAll(".ribbon-help-update-badge").forEach((badge) => {
    badge.classList.toggle("hidden", !shouldSuggestRuntimeDownload);
  });
}



function closeHelpDialog(options = {}) {
  const previousFocus = state.helpDialog?.previousFocus;
  document.querySelector("#helpDialogOverlay")?.remove();
  state.helpDialog = null;
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) {
    previousFocus.focus?.({ preventScroll: true });
  }
}



function openHelpDialog() {
  closeFileSideMenu({ restoreHome: false });
  closeCellContextMenu();
  closeSheetTabContextMenu();
  closeHelpDialog({ restoreFocus: false });
  const overlay = document.createElement("div");
  overlay.id = "helpDialogOverlay";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = `
    <div class="excel-mini-dialog help-dialog" role="dialog" aria-modal="true" aria-labelledby="helpDialogTitle">
      <div class="excel-mini-titlebar" id="helpDialogTitle">ヘルプ</div>
      <div class="excel-mini-body help-dialog-body">
        <p class="help-dialog-lead">${escapeHtml(APP_DISPLAY_NAME)} は HTML ネイティブのワークシートです。保存形式の基本は HTML ファイルで、Excel 形式のファイルも読み込み・保存に対応しています。</p>
        <ul class="help-dialog-list">
          <li><strong>ファイルを開く:</strong> 「ファイル」&gt;「開く」、または開始画面に .html/.xlsx/.xlsm/.xls をドラッグします。</li>
          <li><strong>セル編集:</strong> クリックで選択、ダブルクリックまたは F2 で編集、Enter で確定、Esc で取り消します。</li>
          <li><strong>セル内改行:</strong> 編集中に Alt + Enter を押すと同じセル内で改行します。</li>
          <li><strong>保存:</strong> Web サービスから開いたファイルは Cht WebSheet の HTML として保存します。ローカル HTML を開いている場合は、その HTML ファイルへ保存します。</li>
          <li><strong>共通ランタイム:</strong> ダウンロードしなくても Web サービス上で使用できます。共通ランタイムをインストールすると、保存済み HTML をオフラインでも使用できます。</li>
        </ul>
      </div>
      <div class="excel-mini-footer">
        <button class="excel-mini-button is-primary" type="button" data-help-dialog-action="close">OK</button>
      </div>
    </div>`;
  document.body.append(overlay);
  state.helpDialog = { previousFocus: document.activeElement };
  overlay.querySelector("[data-help-dialog-action='close']")?.focus({ preventScroll: true });
}



function handleHelpDialogAction(event) {
  if (!event.currentTarget.closest("#helpDialogOverlay")) return;
  event.preventDefault();
  closeHelpDialog();
}



function handleHelpDialogKeydown(event) {
  if (event.key !== "Escape") return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  closeHelpDialog();
}



function closeVersionDialog(options = {}) {
  const previousFocus = state.versionDialog?.previousFocus;
  document.querySelector("#versionDialogOverlay")?.remove();
  state.versionDialog = null;
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) {
    previousFocus.focus?.({ preventScroll: true });
  }
}



function openVersionDialog() {
  closeFileSideMenu({ restoreHome: false });
  closeCellContextMenu();
  closeSheetTabContextMenu();
  closeVersionDialog({ restoreFocus: false });
  const overlay = document.createElement("div");
  overlay.id = "versionDialogOverlay";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = `
    <div class="excel-mini-dialog version-dialog" role="dialog" aria-modal="true" aria-labelledby="versionDialogTitle">
      <div class="excel-mini-titlebar" id="versionDialogTitle">バージョン情報</div>
      <div class="excel-mini-body version-dialog-body" data-version-dialog-body></div>
      <div class="excel-mini-footer version-dialog-footer" data-version-dialog-footer></div>
    </div>`;
  document.body.append(overlay);
  state.versionDialog = { previousFocus: document.activeElement };
  renderVersionDialogContent();
  overlay.querySelector("[data-version-dialog-action='close']")?.focus({ preventScroll: true });
  if (state.serverVersionStatus !== "loaded" && state.serverVersionStatus !== "loading") {
    void refreshServerVersionInfo({ force: true });
  }
  void refreshLocalRuntimeInfo({ force: true });
}



function renderVersionDialogContent() {
  const overlay = document.querySelector("#versionDialogOverlay");
  if (!overlay) return;
  const body = overlay.querySelector("[data-version-dialog-body]");
  const footer = overlay.querySelector("[data-version-dialog-footer]");
  if (!body || !footer) return;
  const showDownloadButton = serverRuntimeIsActive() || state.versionMismatch || versionDialogShouldSuggestDownload();
  body.innerHTML = `
    <div class="version-dialog-product">
      <div class="version-dialog-logo" aria-hidden="true">${escapeHtml(APP_SHORT_NAME)}</div>
      <div>
        <div class="version-dialog-name">${escapeHtml(APP_DISPLAY_NAME)}</div>
        <div class="version-dialog-company">${escapeHtml(APP_COMPANY_NAME)}</div>
      </div>
    </div>
    <div class="version-dialog-list">
      <div class="version-dialog-row">
        <div class="version-dialog-label">サーバー バージョン</div>
        <div class="version-dialog-value">${escapeHtml(serverVersionLabel())}</div>
        <div class="version-dialog-updated">${escapeHtml(serverRuntimeUpdatedAtLabel())}</div>
      </div>
      <div class="version-dialog-row">
        <div class="version-dialog-label">ローカル バージョン</div>
        <div class="version-dialog-value">${escapeHtml(localRuntimeStatusLabel())}</div>
        <div class="version-dialog-updated">${escapeHtml(localRuntimeUpdatedAtLabel())}</div>
      </div>
    </div>
    ${versionDialogBenefitHtml()}
    ${versionDialogNoticeHtml()}
  `;
  footer.innerHTML = `
    <button class="excel-mini-button" type="button" data-version-dialog-action="refresh">再確認</button>
    ${showDownloadButton ? '<button class="excel-mini-button is-primary" type="button" data-version-dialog-action="download">ダウンロード</button>' : ""}
    <button class="excel-mini-button${showDownloadButton ? "" : " is-primary"}" type="button" data-version-dialog-action="close">閉じる</button>
  `;
  createIcons({ icons });
}



function runtimeExecutionSource() {
  const source = String(window.__WEBSHEET_RUNTIME_SOURCE__ || "").toLowerCase();
  if (source === "local" || source === "remote") return source;
  if (window.__WEBSHEET_STANDALONE__) return "standalone";
  return "web";
}



function serverVersionLabel() {
  const label = state.serverVersionStatus === "loaded"
    ? serverRuntimeVersionLabel()
    : versionStatusLabel();
  return serverRuntimeIsActive() ? `${label}（使用中）` : label;
}



function serverRuntimeVersionLabel() {
  if (state.serverVersionStatus === "loaded") {
    return versionInfoRuntimeVersion(state.serverVersionInfo) || currentRuntimeVersion() || "-";
  }
  return currentRuntimeVersion() || "-";
}



function serverRuntimeUpdatedAtLabel() {
  if (state.serverVersionStatus === "loading") return "確認中...";
  if (state.serverVersionStatus !== "loaded") return "-";
  return formatVersionUpdatedAt(state.serverVersionInfo?.runtimeUpdatedAt || state.serverVersionInfo?.updatedAt) || "-";
}



function localRuntimeStatusLabel() {
  const source = runtimeExecutionSource();
  const status = localRuntimeStatus();
  const localVersion = localRuntimeDetectedVersion();
  if (source === "local") return `${localVersion || currentRuntimeVersion()}（使用中）`;
  if (status === "checking") return "確認中...";
  if (localRuntimeProbeIsBlockedByBrowser()) return localVersion ? `${localVersion}（確認制限あり）` : "確認できません";
  if (status === "missing") return "未インストール";
  if (status === "outdated") return localVersion ? `${localVersion}（旧バージョン）` : "旧バージョン";
  if (status === "current") return `${localVersion || currentRuntimeVersion()}（インストール済み）`;
  if (source === "remote") return "未インストール";
  if (source === "standalone") return "確認できません";
  return "Web版からは確認できません";
}



function localRuntimeUpdatedAtLabel() {
  const source = runtimeExecutionSource();
  const status = localRuntimeStatus();
  if (source === "local" || status === "current" || status === "outdated") {
    return formatVersionUpdatedAt(localRuntimeDetectedUpdatedAt()) || "-";
  }
  if (status === "checking") return "確認中...";
  return "-";
}



function serverRuntimeIsActive() {
  const source = runtimeExecutionSource();
  return source === "web" || source === "remote";
}



function localRuntimeStatus() {
  return String(window.__WEBSHEET_LOCAL_RUNTIME_STATUS__ || "").toLowerCase();
}



function localRuntimeDetectedVersion() {
  const info = window.__WEBSHEET_LOCAL_RUNTIME_INFO__;
  return versionInfoRuntimeVersion(info);
}



function localRuntimeDetectedUpdatedAt() {
  const info = window.__WEBSHEET_LOCAL_RUNTIME_INFO__;
  return versionInfoRuntimeUpdatedAt(info);
}



function formatVersionUpdatedAt(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const date = new Date(text);
  if (!Number.isFinite(date.getTime())) return text;
  try {
    return new Intl.DateTimeFormat("ja-JP", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    }).format(date);
  } catch {
    return text;
  }
}



function localRuntimeProbeIsBlockedByBrowser() {
  if (localRuntimeStatus() !== "missing") return false;
  if (runtimeExecutionSource() !== "remote") return false;
  try {
    if (!["blob:", "http:", "https:"].includes(window.location.protocol)) return false;
    return (window.__WEBSHEET_LOCAL_RUNTIME_CANDIDATES__ || []).some((candidate) => /^file:/i.test(String(candidate || "")));
  } catch {
    return false;
  }
}



function versionDialogShouldSuggestDownload() {
  const status = localRuntimeStatus();
  if (status === "checking") return false;
  if (localRuntimeProbeIsBlockedByBrowser()) return false;
  return status === "missing" || status === "outdated";
}



function versionStatusLabel() {
  if (state.serverVersionStatus === "loading") return "確認中...";
  if (state.serverVersionStatus === "error") return "確認できませんでした";
  return "未確認";
}



function versionDialogBenefitHtml() {
  if (runtimeExecutionSource() === "local") return "";
  return `
    <div class="version-dialog-benefit">
      <span>ローカル版をインストールすると、保存した CWS HTML ファイルをオフラインでも開けます。</span>
      <span>共通 runtime を毎回ダウンロードしないため、表示開始も安定します。</span>
    </div>`;
}



function versionDialogNoticeHtml() {
  if (state.serverVersionStatus === "loading") {
    return '<div class="version-dialog-notice">サーバー バージョンを確認しています...</div>';
  }
  const localStatus = localRuntimeStatus();
  if (localStatus === "checking") {
    return '<div class="version-dialog-notice">ローカル バージョンを確認しています...</div>';
  }
  if (localRuntimeProbeIsBlockedByBrowser()) {
    return '<div class="version-dialog-notice is-warning">ブラウザーの制限により、この画面からローカル バージョンを確認できません。保存済み CWS HTML をダブルクリックで直接開くと、ローカル runtime の使用状態を確認できます。</div>';
  }
  if (localStatus === "missing") {
    return '<div class="version-dialog-notice is-warning">ローカル バージョンが見つからなかったため、サーバー版で起動しています。インストールするとオフラインでも使用できます。</div>';
  }
  if (localStatus === "outdated") {
    const localVersion = localRuntimeDetectedVersion();
    return `<div class="version-dialog-notice is-warning">ローカル バージョン${localVersion ? ` ${escapeHtml(localVersion)}` : ""} はサーバー版より古いため、サーバー版で起動しています。更新するとオフラインでも使用できます。</div>`;
  }
  if (localStatus === "current" && runtimeExecutionSource() === "remote") {
    return '<div class="version-dialog-notice is-current">ローカル バージョンはインストール済みです。この HTML はサーバー版で起動しましたが、再保存すると次回からローカル版を優先して使用できます。</div>';
  }
  if (state.serverVersionStatus === "error") {
    return '<div class="version-dialog-notice is-warning">サーバー バージョンを確認できませんでした。ネットワークまたは dev server を確認してください。</div>';
  }
  if (!state.versionMismatch) {
    if (runtimeExecutionSource() === "web") {
      return `
        <div class="version-dialog-notice">
          <span>Web版はサーバー版で動作しています。ローカル バージョンのインストール状態は、ブラウザーの制限によりこの画面から確認できません。</span>
          <span>確認するには、現在のブックを CWS HTML として保存し、保存した HTML ファイルをダブルクリックで開いてください。</span>
          <span>その後、ヘルプのバージョン情報でローカル バージョンが使用中か確認できます。</span>
        </div>`;
    }
    return '<div class="version-dialog-notice is-current">ローカル版で起動しています。オフラインでも使用できます。</div>';
  }
  const serverInfo = state.serverVersionInfo;
  return `
    <div class="version-dialog-notice is-update">
      <strong>新しいバージョンがあります。</strong>
      <span>サーバー runtime: ${escapeHtml(versionInfoRuntimeVersion(serverInfo) || "-")}、現在 runtime: ${escapeHtml(currentRuntimeVersion() || "-")}</span>
      <span>ダウンロードページから OS に合うインストーラーを取得し、ローカルの minify 済み JS/CSS runtime を更新してください。</span>
    </div>`;
}



function handleVersionDialogAction(event) {
  const action = event.currentTarget?.dataset?.versionDialogAction;
  if (!action || !event.currentTarget.closest("#versionDialogOverlay")) return;
  event.preventDefault();
  if (action === "close") {
    closeVersionDialog();
    return;
  }
  if (action === "refresh") {
    void refreshServerVersionInfo({ force: true });
    void refreshLocalRuntimeInfo({ force: true });
    return;
  }
  if (action === "download") {
    openVersionDownloadPage();
  }
}



function openVersionDownloadPage() {
  const downloadPageUrl = versionDownloadPageUrl();
  const openedWindow = window.open(downloadPageUrl, "_blank", "noopener,noreferrer");
  if (openedWindow) {
    openedWindow.opener = null;
  }
}



function handleVersionDialogKeydown(event) {
  if (event.key !== "Escape") return;
  if (!document.querySelector("#versionDialogOverlay")) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  event.stopPropagation();
  closeVersionDialog();
}



function versionDownloadPageUrl() {
  const configured = state.serverVersionInfo?.downloadPage || SERVER_DOWNLOAD_PAGE_PATH;
  try {
    if (/^https?:/i.test(configured)) return configured;
    const versionInfoUrl = new URL(serverVersionInfoUrl(), window.location.href);
    const appBaseUrl = versionInfoUrl.href.replace(/download\/version\.json(?:[?#].*)?$/i, "");
    return new URL(configured.replace(/^\/+/, ""), appBaseUrl).href;
  } catch {
    return configured;
  }
}



function showAppAlert(message, tone = "info") {
  $appAlert
    .removeClass("hidden is-error is-info")
    .addClass(tone === "error" ? "is-error" : "is-info")
    .text(message);
}



function clearAppAlert() {
  $appAlert.addClass("hidden").removeClass("is-error is-info").empty();
}



function showMessageBox(message, options = {}) {
  closeMessageBox();
  const title = options.title || "WebSheet";
  const tone = ["stop", "warning", "info"].includes(options.tone) ? options.tone : "info";
  const icon = options.icon || (tone === "stop" ? "×" : tone === "info" ? "i" : "!");
  const buttons = Array.isArray(options.buttons) && options.buttons.length ? options.buttons : [{ label: "OK", result: "ok" }];
  const callbacks = {};
  buttons.forEach((button, index) => {
    const result = button.result || `button-${index}`;
    if (typeof button.callback === "function") callbacks[result] = button.callback;
  });
  state.messageBoxHandlers = {
    callbacks,
    cancelResult: options.cancelResult || buttons[buttons.length - 1]?.result || "ok",
  };
  const overlay = document.createElement("div");
  overlay.id = "messageBoxOverlay";
  overlay.className = "message-box-overlay";
  overlay.innerHTML = `
    <div id="messageBox" class="message-box" role="alertdialog" aria-modal="true" aria-labelledby="messageBoxTitle" aria-describedby="messageBoxText">
      <div class="message-box-title" id="messageBoxTitle">${escapeHtml(title)}</div>
      <div class="message-box-body">
        <div class="message-box-icon is-${escapeAttr(tone)}" aria-hidden="true">${escapeHtml(icon)}</div>
        <div class="message-box-text" id="messageBoxText">${escapeHtml(message)}</div>
      </div>
      <div class="message-box-actions">
        ${buttons
          .map((button, index) => {
            const result = button.result || `button-${index}`;
            const label = button.label || "OK";
            const accessKey = menuAccessKeyFromLabel(label);
            const accessKeyAttr = accessKey ? ` data-access-key="${escapeAttr(accessKey)}"` : "";
            return `<button class="message-box-button" type="button" data-message-box-close data-message-box-result="${escapeAttr(result)}"${accessKeyAttr}>${accessKeyLabelHtml(label)}</button>`;
          })
          .join("")}
      </div>
    </div>
  `;
  document.body.append(overlay);
  overlay.querySelector("[data-message-box-close]")?.focus({ preventScroll: true });
}



function handleMessageBoxButton(event) {
  const result = event.currentTarget?.dataset?.messageBoxResult || "ok";
  closeMessageBox(result, true);
}



function handleMessageBoxKeydown(event) {
  if (isElementImeCompositionEvent(event)) return;
  const overlay = event.currentTarget;
  if (event.key === "Enter") {
    const target = document.activeElement?.closest?.("[data-message-box-close]") ||
      overlay.querySelector("[data-message-box-close]");
    if (!target) return;
    event.preventDefault();
    target.click();
    return;
  }
  const accessKey = accessKeyFromKeyboardEvent(event);
  if (!accessKey || event.ctrlKey || event.metaKey || event.altKey) return;
  const target = [...overlay.querySelectorAll("[data-message-box-close]")]
    .find((button) => String(button.dataset.accessKey || "").toUpperCase() === accessKey);
  if (!target) return;
  event.preventDefault();
  target.click();
}



function closeMessageBox(result = null, invokeCallback = false) {
  const handlers = state.messageBoxHandlers;
  document.querySelector("#messageBoxOverlay")?.remove();
  state.messageBoxHandlers = null;
  if (!invokeCallback || !handlers) return;
  const key = result || handlers.cancelResult;
  const callback = handlers.callbacks?.[key];
  if (callback) callback();
}



function closeMessageBoxOnEscape(event) {
  if (event.key !== "Escape" || !document.querySelector("#messageBoxOverlay")) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  closeMessageBox(state.messageBoxHandlers?.cancelResult || "cancel", true);
}



function showWarnings(messages) {
  const unique = [...new Set(messages.filter(Boolean))];
  if (!unique.length) return clearWarnings();
  $warnings.removeClass("hidden").html(unique.map((message) => `<div>${escapeHtml(message)}</div>`).join(""));
}



function clearWarnings() {
  $warnings.addClass("hidden").empty();
}



function readableLoadError(error) {
  const rawMessage = String(error?.message || error || "不明なエラー");
  if (/HTML workbook data does not contain sheets/i.test(rawMessage)) {
    return "HTML 内のワークブック データにワークシート情報がありません。";
  }
  if (/Unexpected end of JSON input|JSON\.parse|unterminated string|bad control character|Expected .*JSON/i.test(rawMessage)) {
    return "HTML 内のワークブック データを解析できません。保存済み HTML が破損している可能性があります。";
  }
  if (/password|encrypted|unsupported.*encryption/i.test(rawMessage)) {
    return "このファイルはパスワード保護または暗号化されている可能性があり、ブラウザーでは読み取れません。暗号化されていない .xlsx として保存し直してからお試しください。";
  }
  if (/Can't find end of central directory|Corrupted zip|invalid signature|zip/i.test(rawMessage)) {
    return "このファイルは有効な .xlsx/.xlsm/.xls ブックではないようです。Excel で開いて .xlsx として保存し直してからアップロードしてください。";
  }
  if (/Cannot read properties|undefined|null/i.test(rawMessage)) {
    return `未対応のブック構造を解析中に検出しました。元のエラー：${rawMessage}`;
  }
  return rawMessage;
}



function columnWidthToPixels(width, scale = 1, maxDigitWidth = 7) {
  const digitWidth = Number.isFinite(Number(maxDigitWidth)) && Number(maxDigitWidth) > 0 ? Number(maxDigitWidth) : 7;
  if (!width) return Math.max(1, Math.round(64 * scale));
  const basePixels = width < 1 ? Math.floor(width * (digitWidth + 5) + 0.5) : Math.floor(width * digitWidth + 5);
  return Math.max(1, Math.round(basePixels * scale));
}



function rowHeightToPixels(height, scale = 1) {
  if (!height) return Math.max(1, Math.round(20 * scale));
  return Math.max(1, Math.round(height * (96 / 72) * scale));
}



function dateToExcelSerial(date) {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds());
  return utc / 86400000 + 25569;
}



function excelDateObjectToSerial(date) {
  const utc = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds());
  return utc / 86400000 + 25569;
}



function workbookUses1904DateSystem(source = state.model) {
  return source?.workbookProperties?.date1904 === true || source?.date1904 === true;
}



function excelSerial1904To1900(serial) {
  const value = Number(serial);
  return Number.isFinite(value) ? value + 1462 : serial;
}



function excelSerial1900To1904(serial) {
  const value = Number(serial);
  return Number.isFinite(value) ? value - 1462 : serial;
}



function excelSerialToDate(serial) {
  const date = new Date(Math.round((serial - 25569) * 86400000));
  return new Date(
    date.getUTCFullYear(),
    date.getUTCMonth(),
    date.getUTCDate(),
    date.getUTCHours(),
    date.getUTCMinutes(),
    date.getUTCSeconds(),
  );
}



function isDateFormat(numFmt) {
  if (isElapsedTimeFormat(numFmt)) return true;
  const format = dateDetectionFormat(numFmt);
  if (/[年月日]/.test(format)) return true;
  if (/^\s*(?:y{1,4}|m{1,4}|d{1,4})\s*$/i.test(format)) return true;
  const hasYear = /(^|[^a-z])y{2,4}/i.test(format);
  const hasMonth = /(^|[^a-z])m{1,4}/i.test(format);
  const hasDay = /(^|[^a-z])d{1,4}/i.test(format);
  const hasHour = /(^|[^a-z])h{1,2}/i.test(format);
  const hasSecond = /(^|[^a-z])s{1,2}/i.test(format);
  const hasAmPm = /AM\/PM|A\/P/i.test(format);
  return (hasYear && hasMonth) || (hasMonth && hasDay) || (hasYear && hasDay) || hasHour || hasSecond || hasAmPm;
}



function dateDetectionFormat(numFmt) {
  return primaryExcelFormatSection(numFmt || "")
    .replace(/\[[^\]]+\]/g, "")
    .replace(/"[^"]*"/g, " ")
    .replace(/\\./g, " ");
}



function currencySymbol(numFmt) {
  const match = String(numFmt || "").match(/[$¥€£]/);
  return match ? match[0] : "";
}



function cellKey(row, col) {
  return `${row}:${col}`;
}



function columnName(index) {
  let name = "";
  let value = index;
  while (value > 0) {
    const modulo = (value - 1) % 26;
    name = String.fromCharCode(65 + modulo) + name;
    value = Math.floor((value - modulo) / 26);
  }
  return name;
}



function columnNumber(name) {
  return name
    .toUpperCase()
    .split("")
    .reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}



function decodeCell(address) {
  const match = String(address).match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) return null;
  return {
    col: columnNumber(match[1]),
    row: Number(match[2]),
  };
}



function decodeRange(range) {
  const [start, end = start] = String(range).split(":");
  const startCell = decodeCell(start);
  const endCell = decodeCell(end);
  if (!startCell || !endCell) return null;
  return {
    top: Math.min(startCell.row, endCell.row),
    left: Math.min(startCell.col, endCell.col),
    bottom: Math.max(startCell.row, endCell.row),
    right: Math.max(startCell.col, endCell.col),
  };
}



function styleObjectToString(style) {
  return Object.entries(style || {})
    .map(([key, value]) => `${key.replace(/[A-Z]/g, (match) => `-${match.toLowerCase()}`)}:${value}`)
    .join(";");
}



function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => {
    const entities = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;",
    };
    return entities[char];
  });
}



function accessKeyLabelHtml(value) {
  const escaped = escapeHtml(value);
  const withParentheses = escaped.replace(/\(([A-Za-z0-9])\)|（([A-Za-z0-9])）/g, (_match, asciiKey, fullWidthKey) => {
    if (asciiKey) return `<span class="menu-access-key-token">(<span class="menu-access-key">${asciiKey}</span>)</span>`;
    return `<span class="menu-access-key-token">（<span class="menu-access-key">${fullWidthKey}</span>）</span>`;
  });
  if (withParentheses !== escaped) return withParentheses;
  return escaped.replace(/^(\s*)([0-9])(\s+)/, (_match, prefix, accessKey, suffix) => `${prefix}<span class="menu-access-key-token"><span class="menu-access-key">${accessKey}</span></span>${suffix}`);
}



function escapeAttr(value) {
  return escapeHtml(value).replace(/`/g, "&#96;");
}



function safeJson(value) {
  return JSON.stringify(value).replace(/</g, "\\u003c").replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029");
}



function sanitizeFileName(value) {
  return value.replace(/[\\/:*?"<>|]+/g, "-").trim();
}

export {
  accessKeyLabelHtml,
  activeRuntimeVersionInfo,
  canProbeLocalRuntimeVersion,
  cellKey,
  clearAppAlert,
  clearWarnings,
  closeDialogFromChrome,
  closeHelpDialog,
  closeMessageBox,
  closeMessageBoxOnEscape,
  closeVersionDialog,
  columnName,
  columnNumber,
  columnWidthToPixels,
  compareRuntimeVersions,
  configuredLocalRuntimeBaseUrls,
  currencySymbol,
  currentRuntimeUpdatedAt,
  currentRuntimeVersion,
  dateDetectionFormat,
  dateToExcelSerial,
  decodeCell,
  decodeRange,
  decorateDialogChrome,
  decorateSingleDialogChrome,
  escapeAttr,
  escapeHtml,
  excelDateObjectToSerial,
  excelSerial1900To1904,
  excelSerial1904To1900,
  excelSerialToDate,
  fetchLocalRuntimeVersionJson,
  finishDialogDrag,
  formatVersionUpdatedAt,
  handleDialogChromeCloseClick,
  handleHelpDialogAction,
  handleHelpDialogKeydown,
  handleMessageBoxButton,
  handleMessageBoxKeydown,
  handleVersionDialogAction,
  handleVersionDialogKeydown,
  inferredUserRuntimeBaseUrlsFromCurrentFile,
  initDialogChromeObserver,
  isDateFormat,
  localRuntimeDetectedUpdatedAt,
  localRuntimeDetectedVersion,
  localRuntimeProbeIsBlockedByBrowser,
  localRuntimeStatus,
  localRuntimeStatusLabel,
  localRuntimeUpdatedAtLabel,
  normalizeVersionInfo,
  openHelpDialog,
  openVersionDialog,
  openVersionDownloadPage,
  probeLocalRuntimeVersion,
  readLocalRuntimeVersion,
  readableLoadError,
  refreshLocalRuntimeInfo,
  refreshServerVersionInfo,
  renderVersionDialogContent,
  requiredRuntimeUpdatedAt,
  requiredRuntimeVersion,
  rowHeightToPixels,
  runtimeExecutionSource,
  runtimeInfoIsUsable,
  runtimeUpdatedAtTimestamp,
  runtimeVersionParts,
  safeJson,
  sanitizeFileName,
  serverRuntimeIsActive,
  serverRuntimeUpdatedAtLabel,
  serverRuntimeVersionLabel,
  serverVersionInfoUrl,
  serverVersionLabel,
  setStatus,
  showAppAlert,
  showMessageBox,
  showWarnings,
  standaloneCss,
  startDialogDrag,
  styleObjectToString,
  syncVersionRibbonBadge,
  updateDialogDrag,
  versionDialogBenefitHtml,
  versionDialogNoticeHtml,
  versionDialogShouldSuggestDownload,
  versionDownloadPageUrl,
  versionInfoAppVersion,
  versionInfoIsMismatch,
  versionInfoPayload,
  versionInfoRuntimeUpdatedAt,
  versionInfoRuntimeVersion,
  versionInfoUpdatedAt,
  versionStatusLabel,
  workbookUses1904DateSystem,
};
