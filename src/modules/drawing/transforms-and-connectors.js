// Feature module; shared bindings are imported explicitly.
import {
  $sheetHost,
  COLUMN_HEADER_HEIGHT,
  CUSTOM_SHAPE_TYPE,
  DEFAULT_SHAPE_TYPE,
  LINE_CONNECTION_RETENTION_TOLERANCE_PX,
  SHAPE_ARROW_PRESETS,
  state,
} from "../../app-core.js";
import {
  chartSvg,
  lineBoxAndPointsFromEndpoints,
} from "./charts-and-insertion.js";
import {
  clamp01,
  clampNumber,
  defaultCalloutTailX,
  defaultCalloutTailY,
  elbowConnectorRoutePointsPx,
  explosionBurstPoints,
  explosionSealPoints,
  isCalloutShapeType,
  isLineShapeObject,
  linePointToPixelPoint,
  normalizeLineShapeConnections,
  normalizeRotationDegrees,
  normalizeShapeAdjustments,
  normalizeShapeArrowSize,
  normalizeShapeArrowType,
  normalizeShapeCustomPath,
  normalizeShapeGlow,
  normalizeShapeLineCompound,
  normalizeShapeLineDash,
  normalizeShapeLinePoints,
  normalizeShapeOutlineWidth,
  normalizeShapeReflection,
  normalizeShapeSoftEdges,
  roundRectCornerScale,
  shapeCanEditPoints,
  shapeDescriptorForType,
  starPointCountForType,
} from "./shape-rendering.js";
import {
  commitActiveCellEditForExternalPointer,
  correspondingGroupedShapeObject,
  groupedShapeObjectTargets,
  openShapeTextEditor,
  selectImage,
  setPointerCaptureSafely,
  startPendingObjectDrawingFromPointer,
} from "../editing/autofill.js";
import {
  normalizeColorValue,
  roundSvgNumber,
  svgDataUrl,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  copiedObjectName,
  selectedObjectSummary,
} from "../editing/keyboard-and-paste.js";
import {
  cssEscape,
  mutateShapeObjects,
  nativeAnchorFromSheetPixel,
  restoreGroupedSheetsAfterFailure,
  sheetIndexFromPixel,
  sheetPixelFromClientPosition,
  syncShapeFormatControls,
  updateSelectionUi,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  clampElbowConnectorBendForPoints,
  clearPictureCropState,
  coordinateToPixels,
  customPathEditablePointCount,
  customPathVertexForKey,
  elbowConnectorUsesHorizontalMiddle,
  imageBoxStyle,
  imageBoxToPixels,
  imageFlipH,
  imageFlipV,
  imageRotationDegrees,
  imageSelectionFrameStyle,
  imageViewportBounds,
  imageVisualBoundsToPixels,
  imageVisualBoxToPixels,
  isActiveImage,
  isPictureCropModeActive,
  pixelBoxesIntersect,
  roundCssNumber,
  selectedCroppablePicture,
  selectedImageGroups,
  selectedImageObjectById,
  selectedImageObjects,
  setImageElementRenderSource,
  setSelectedImageIds,
  shapeAdjustmentBoxFromFrame,
  shapeAdjustmentHandlesForFrame,
  shapeCustomPathHandles,
  shapeCustomPathIsClosed,
  unionImageBox,
  unionImageSelectionBounds,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneImages,
  pushHistory,
} from "../editing/structure-and-history.js";
import {
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  setStatus,
} from "../ui/help-and-utilities.js";
import {
  ensureSheetSize,
  rowHeaderWidthForSheet,
} from "../ui/ribbon-and-links.js";
import {
  ensureSheetProtectionPermission,
} from "../ui/sheet-protection-and-translation.js";
import {
  normalizePictureCrop,
} from "../workbook/ooxml-drawings.js";
import {
  activeSheet,
  visibleGridRange,
} from "../workbook/saving-and-export.js";


function moveSelectedImagesBySheetGroup(objects, dx, dy) {
  const items = groupedShapeObjectTargets(objects);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを移動できません。")) return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      const movedShapeIds = [];
      const movedLineIds = [];
      item.images.forEach((image) => {
        const box = imageBoxToPixels(item.sheet, image);
        if (!box) return;
        setImagePixelBox(item.sheet, image, {
          ...box,
          left: box.left + dx,
          top: box.top + dy,
        });
        if (isLineShapeObject(image)) movedLineIds.push(image.id);
        else movedShapeIds.push(image.id);
      });
      updateConnectedLinesForTargets(item.sheet, movedShapeIds);
      pruneDetachedLineShapeConnectionsForIds(item.sheet, movedLineIds);
    });
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート${selectedObjectSummary(objects)} 移動`,
    });
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
    setStatus(`${items.length} 個のシートの${selectedObjectSummary(objects)}を移動しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクトを移動できませんでした：${error.message || error}`);
  }
  return true;
}



function togglePictureCropMode() {
  const image = selectedCroppablePicture();
  if (!image) {
    setStatus("トリミングする画像を 1 つ選択してください。");
    return false;
  }
  if (isPictureCropModeActive(image)) {
    clearPictureCropState();
    setStatus("トリミングを終了しました。");
  } else {
    state.pictureCrop = { sheetIndex: state.activeSheetIndex, imageId: image.id };
    setStatus("黒いトリミング ハンドルをドラッグして画像を切り抜けます。");
  }
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  return true;
}



function currentPictureCropForImage(image) {
  return normalizePictureCrop(image?.crop) || { left: 0, top: 0, right: 0, bottom: 0 };
}



function pictureCropFullBoxFromVisibleBox(box, crop = {}) {
  const normalized = crop || { left: 0, top: 0, right: 0, bottom: 0 };
  const visibleWidth = Math.max(0.02, 1 - normalized.left - normalized.right);
  const visibleHeight = Math.max(0.02, 1 - normalized.top - normalized.bottom);
  const width = Math.max(1, box.width / visibleWidth);
  const height = Math.max(1, box.height / visibleHeight);
  return {
    left: box.left - width * normalized.left,
    top: box.top - height * normalized.top,
    width,
    height,
  };
}



function pictureCropVisibleBoxFromFullBox(fullBox, crop = {}) {
  const normalized = crop || { left: 0, top: 0, right: 0, bottom: 0 };
  const visibleWidth = Math.max(0.02, 1 - normalized.left - normalized.right);
  const visibleHeight = Math.max(0.02, 1 - normalized.top - normalized.bottom);
  return {
    left: fullBox.left + fullBox.width * normalized.left,
    top: fullBox.top + fullBox.height * normalized.top,
    width: Math.max(1, fullBox.width * visibleWidth),
    height: Math.max(1, fullBox.height * visibleHeight),
  };
}



function normalizeInteractivePictureCrop(crop = {}) {
  let left = clampNumber(crop.left, 0, 0.98, 0);
  let top = clampNumber(crop.top, 0, 0.98, 0);
  let right = clampNumber(crop.right, 0, 0.98, 0);
  let bottom = clampNumber(crop.bottom, 0, 0.98, 0);
  const horizontal = left + right;
  if (horizontal > 0.98) {
    const scale = 0.98 / horizontal;
    left *= scale;
    right *= scale;
  }
  const vertical = top + bottom;
  if (vertical > 0.98) {
    const scale = 0.98 / vertical;
    top *= scale;
    bottom *= scale;
  }
  const rounded = {
    left: Math.round(left * 10000) / 10000,
    top: Math.round(top * 10000) / 10000,
    right: Math.round(right * 10000) / 10000,
    bottom: Math.round(bottom * 10000) / 10000,
  };
  return rounded.left || rounded.top || rounded.right || rounded.bottom ? rounded : null;
}



function startPictureCropDrag(event) {
  if (event.button !== 0) return;
  const handle = event.currentTarget?.dataset?.pictureCropHandle || "";
  const image = selectedCroppablePicture();
  if (!handle || !image || !isPictureCropModeActive(image)) return;
  if (!ensureSheetProtectionPermission(activeSheet(), "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return;
  const sheet = activeSheet();
  const box = imageBoxToPixels(sheet, image);
  if (!box) return;
  const startCrop = currentPictureCropForImage(image);
  state.pictureCropDrag = {
    sheetIndex: state.activeSheetIndex,
    imageId: image.id,
    handle,
    startClientX: event.clientX,
    startClientY: event.clientY,
    startCrop,
    box: { ...box },
    fullBox: pictureCropFullBoxFromVisibleBox(box, startCrop),
    sheetBefore: captureSheetForHistory(sheet),
    pointerId: event.pointerId,
    changed: false,
  };
  setPointerCaptureSafely(event.currentTarget, event.pointerId);
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
}



function updatePictureCropDrag(event) {
  const job = state.pictureCropDrag;
  if (!job) return;
  if (job.pointerId && event.pointerId && job.pointerId !== event.pointerId) return;
  const sheet = activeSheet();
  const image = sheet?.images?.find((item) => item.id === job.imageId);
  if (!image) return;
  const fullBox = job.fullBox || pictureCropFullBoxFromVisibleBox(job.box, job.startCrop);
  const dx = (event.clientX - job.startClientX) / Math.max(1, fullBox.width);
  const dy = (event.clientY - job.startClientY) / Math.max(1, fullBox.height);
  const next = { ...job.startCrop };
  if (job.handle.includes("w")) next.left = job.startCrop.left + dx;
  if (job.handle.includes("e")) next.right = job.startCrop.right - dx;
  if (job.handle.includes("n")) next.top = job.startCrop.top + dy;
  if (job.handle.includes("s")) next.bottom = job.startCrop.bottom - dy;
  image.crop = normalizeInteractivePictureCrop(next);
  setImagePixelBox(sheet, image, pictureCropVisibleBoxFromFullBox(fullBox, image.crop || { left: 0, top: 0, right: 0, bottom: 0 }));
  job.changed = true;
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  event.preventDefault?.();
}



function finishPictureCropDrag(event) {
  const job = state.pictureCropDrag;
  if (!job) return;
  if (job.pointerId && event?.pointerId && job.pointerId !== event.pointerId) return;
  state.pictureCropDrag = null;
  if (!job.changed) return;
  const sheet = activeSheet();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore: job.sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: "画像 トリミング",
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus("画像をトリミングしました。");
  event?.preventDefault?.();
}



function enterShapePointEditMode() {
  const objects = selectedImageObjects().filter(shapeCanEditPoints);
  if (!objects.length) {
    setStatus("頂点を編集できる図形を選択してください。");
    return false;
  }
  state.shapePointEditIds = objects.map((image) => image.id);
  renderSheet();
  updateSelectionUi();
  const lineCount = objects.filter(isLineShapeObject).length;
  const message = lineCount
    ? "端点をドラッグして線を調整できます。"
    : "黄色の調整ハンドルをドラッグして図形を調整できます。";
  setStatus(message);
  return true;
}



function moveSelectedImagesBy(dx, dy) {
  const objects = selectedImageObjects();
  if (!objects.length || (!dx && !dy)) return false;
  const sheet = activeSheet();
  if (moveSelectedImagesBySheetGroup(objects, dx, dy)) return true;
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを移動できません。")) return false;
  const sheetBefore = captureSheetForHistory(sheet);
  const movedLineIds = [];
  const movedShapeIds = [];
  objects.forEach((image) => {
    const box = imageBoxToPixels(sheet, image);
    if (!box) return;
    setImagePixelBox(sheet, image, {
      ...box,
      left: box.left + dx,
      top: box.top + dy,
    });
    if (isLineShapeObject(image)) movedLineIds.push(image.id);
    else movedShapeIds.push(image.id);
  });
  updateConnectedLinesForTargets(sheet, movedShapeIds);
  pruneDetachedLineShapeConnectionsForIds(sheet, movedLineIds);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${selectedObjectSummary(objects)} 移動`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`${selectedObjectSummary(objects)}を移動しました。`);
  return true;
}



function selectionIdsForImage(sheet, imageId, options = {}) {
  const image = sheet.images?.find((item) => item.id === imageId);
  if (!image) return [];
  if (options.forceSingle || !image.groupId) return [imageId];
  const groupIds = (sheet.images || [])
    .filter((item) => item.groupId === image.groupId)
    .map((item) => item.id);
  return groupIds.length > 1 ? groupIds : [imageId];
}



function startSelectedImageTransform(event, options = {}) {
  if (event.button !== 0) return;
  if (event.target?.closest?.(".sheet-picture-crop-handle")) return;
  if (startPendingObjectDrawingFromPointer(event)) return;
  commitActiveCellEditForExternalPointer(event);
  if (!ensureSheetProtectionPermission(activeSheet(), "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return;
  const imageId = options.imageId || event.currentTarget?.dataset?.imageId || event.target.closest("[data-image-id]")?.dataset.imageId || state.selectedImage?.id;
  const groupId = options.groupId || event.currentTarget?.dataset?.imageGroupId || "";
  const isPlainShapeDoubleClick =
    event.detail >= 2 &&
    !event.target?.closest?.(".sheet-image-resize-handle, .sheet-shape-adjust-handle, .sheet-image-rotate-handle") &&
    !isLineShapeObject(selectedImageObjectById(imageId)) &&
    selectedImageObjectById(imageId)?.shape;
  if (isPlainShapeDoubleClick) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    openShapeTextEditor(imageId, { clientX: event.clientX, clientY: event.clientY });
    return;
  }
  if (!isActiveImage(imageId)) {
    selectImage(imageId);
  } else if (state.selectedImage?.id !== imageId) {
    state.selectedImage = { sheetIndex: state.activeSheetIndex, id: imageId };
    updateSelectionUi();
  }
  const image = selectedImageObjectById(imageId);
  if (!image) return;
  const sheet = activeSheet();
  const startBox = imageBoxToPixels(sheet, image);
  const pointer = sheetPixelFromClientPosition(event.clientX, event.clientY, { clampToGrid: true });
  if (!startBox || !pointer) return;

  const adjustKey = event.target?.dataset?.shapeAdjust || "";
  const handle = adjustKey || event.target?.dataset?.imageResize || "";
  const isRotate = Boolean(event.target?.closest?.("[data-image-rotate]"));
  const toggleSelectionOnClick = Boolean((event.ctrlKey || event.metaKey) && !adjustKey && !handle && !isRotate && isActiveImage(image.id));
  const groupMembers = groupId ? selectedImageObjects().filter((item) => item.groupId === groupId) : [];
  const startGroupBox = groupMembers.length > 1 ? unionImageSelectionBounds(sheet, groupMembers) : null;
  const mode = groupId && groupMembers.length > 1 && isRotate
    ? "group-rotate"
    : groupId && handle && !isRotate
      ? "group-resize"
      : adjustKey
        ? "shape-adjust"
        : isRotate
          ? "rotate"
          : handle.startsWith("line-")
            ? "line-point"
            : handle
              ? "resize"
              : "move";
  const transformImages = ["move", "group-resize", "group-rotate"].includes(mode) ? selectedImageObjects() : [image];
  const startBoxes = Object.fromEntries(
    transformImages
      .map((item) => [item.id, imageBoxToPixels(sheet, item)])
      .filter((entry) => entry[1]),
  );
  const startRotations = Object.fromEntries(transformImages.map((item) => [item.id, imageRotationDegrees(item)]));
  state.imageTransform = {
    imageId: image.id,
    imageIds: Object.keys(startBoxes),
    mode,
    handle,
    startPointer: pointer,
    currentBox: { ...startBox },
    currentBoxes: Object.fromEntries(Object.entries(startBoxes).map(([id, box]) => [id, { ...box }])),
    startBox: { ...startBox },
    startBoxes,
    groupId,
    startGroupBox: startGroupBox ? { ...startGroupBox } : null,
    currentGroupBox: startGroupBox ? { ...startGroupBox } : null,
    startRotation: imageRotationDegrees(image),
    currentRotation: imageRotationDegrees(image),
    startRotations,
    currentRotations: { ...startRotations },
    startLinePoints: isLineShapeObject(image) ? normalizeShapeLinePoints(image.shape.points) : null,
    currentLinePoints: isLineShapeObject(image) ? normalizeShapeLinePoints(image.shape.points) : null,
    shapeType: image.shape?.type || "",
    startAdjust: image.shape ? normalizeShapeAdjustments(image.shape.type, image.shape.adjust) : null,
    currentAdjust: image.shape ? normalizeShapeAdjustments(image.shape.type, image.shape.adjust) : null,
    startCustomPath: image.shape ? normalizeShapeCustomPath(image.shape.customPath) : null,
    currentCustomPath: image.shape ? normalizeShapeCustomPath(image.shape.customPath) : null,
    preserveAspectOnCornerResize: shouldPreservePictureAspectOnResize(image, mode, handle),
    toggleSelectionOnClick,
    sheetBefore: captureSheetForHistory(sheet),
    copyOnMove: mode === "move" && Boolean(event.ctrlKey),
    constrainMove: false,
    snapToGrid: false,
    pointerId: event.pointerId,
    changed: false,
  };
  setPointerCaptureSafely(options.captureTarget || event.currentTarget, event.pointerId);
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
}



function updateSelectedImageTransform(event) {
  const job = state.imageTransform;
  if (!job) return;
  if (job.pointerId && event.pointerId && job.pointerId !== event.pointerId) return;
  const pointer = sheetPixelFromClientPosition(event.clientX, event.clientY, { clampToGrid: true });
  if (!pointer) return;
  job.lockAspect = Boolean(event.shiftKey || job.preserveAspectOnCornerResize);
  job.constrainMove = job.mode === "move" && event.shiftKey;
  job.snapToGrid = job.mode === "move" && event.altKey;
  job.copyOnMove = job.mode === "move" && event.ctrlKey;
  if (job.mode === "rotate") {
    job.currentRotation = imageTransformRotation(job, pointer);
  } else if (job.mode === "line-point") {
    const result = imageTransformLinePoint(job, pointer);
    job.currentBox = result.box;
    job.currentLinePoints = result.points;
    job.currentLineConnection = result.connection;
    updateLineConnectionPreviewOverlay(result.connection, job.imageId, result.previewTargetId);
  } else if (job.mode === "shape-adjust") {
    if (job.handle?.startsWith("custom-point:")) {
      job.currentCustomPath = imageTransformCustomPathPoint(job, pointer);
    } else {
      job.currentAdjust = imageTransformShapeAdjustment(job, pointer);
    }
  } else if (job.mode === "group-resize") {
    const result = imageTransformGroupResizeBoxes(job, pointer);
    job.currentBoxes = result.boxes;
    job.currentGroupBox = result.groupBox;
    job.currentBox = result.groupBox;
  } else if (job.mode === "group-rotate") {
    const result = imageTransformGroupRotationBoxes(job, pointer);
    job.currentBoxes = result.boxes;
    job.currentRotations = result.rotations;
    job.currentGroupBox = result.groupBox;
    job.currentBox = result.groupBox;
  } else if (job.mode === "move" && (job.imageIds || []).length > 1) {
    job.currentBoxes = imageTransformMoveBoxes(job, pointer);
    job.currentBox = job.currentBoxes[job.imageId] || job.currentBox;
  } else {
    job.currentBox = imageTransformBox(job, pointer);
  }
  job.changed = true;
  if (job.currentBoxes && Object.keys(job.currentBoxes).length > 1) {
    Object.entries(job.currentBoxes).forEach(([imageId, box]) => previewSelectedImageBox(imageId, box, {
      rotation: job.currentRotations?.[imageId],
    }));
    if (job.groupId) previewSelectedImageGroupFrame(job.groupId, job.currentBoxes, job.currentRotations);
  } else {
    previewSelectedImageBox(job.imageId, job.currentBox, {
      rotation: job.currentRotation,
      points: job.currentLinePoints,
      adjust: job.currentAdjust,
      customPath: job.currentCustomPath,
    });
  }
  previewConnectedLinesForTransform(job);
  event.preventDefault?.();
}



function finishSelectedImageTransform(event) {
  const job = state.imageTransform;
  if (!job) return;
  if (job.pointerId && event.pointerId && job.pointerId !== event.pointerId) return;
  state.imageTransform = null;
  removeLineConnectionPreviewOverlay();
  const rotationChanged = job.mode === "rotate" && !sameRotation(job.startRotation, job.currentRotation);
  const lineChanged = job.mode === "line-point" && (!boxesEqual(job.startBox, job.currentBox) || !linePointsEqual(job.startLinePoints, job.currentLinePoints));
  const adjustChanged = job.mode === "shape-adjust" && !shapeAdjustmentsEqual(job.shapeType, job.startAdjust, job.currentAdjust);
  const customPathChanged = job.mode === "shape-adjust" && !shapeCustomPathsEqual(job.startCustomPath, job.currentCustomPath);
  const groupResizeChanged =
    job.mode === "group-resize" &&
    (job.imageIds || []).some((imageId) => !boxesEqual(job.startBoxes?.[imageId], job.currentBoxes?.[imageId]));
  const groupRotateChanged =
    job.mode === "group-rotate" &&
    (job.imageIds || []).some(
      (imageId) =>
        !boxesEqual(job.startBoxes?.[imageId], job.currentBoxes?.[imageId]) ||
        !sameRotation(job.startRotations?.[imageId], job.currentRotations?.[imageId]),
    );
  const multiMoveChanged =
    job.mode === "move" &&
    (job.imageIds || []).length > 1 &&
    (job.imageIds || []).some((imageId) => !boxesEqual(job.startBoxes?.[imageId], job.currentBoxes?.[imageId]));
  const boxChanged = !["rotate", "line-point", "shape-adjust", "group-resize", "group-rotate"].includes(job.mode) && (multiMoveChanged || !boxesEqual(job.startBox, job.currentBox));
  if (!job.changed || (!rotationChanged && !lineChanged && !adjustChanged && !customPathChanged && !groupResizeChanged && !groupRotateChanged && !boxChanged)) {
    if (job.toggleSelectionOnClick && !job.changed) {
      selectImage(job.imageId, { additive: true, toggle: true });
      event.preventDefault?.();
      return;
    }
    updateSelectionUi();
    return;
  }
  const sheet = activeSheet();
  const image = sheet.images?.find((item) => item.id === job.imageId);
  if (!image) return;
  if (job.copyOnMove && job.mode === "move") {
    const inserted = insertCopiedImagesFromMoveTransform(sheet, job);
    if (!inserted.length) {
      updateSelectionUi();
      return;
    }
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetBefore: job.sheetBefore,
      sheetAfter: captureSheetForHistory(sheet),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${selectedObjectSummary(inserted)} コピー`,
    });
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
    setStatus(`${selectedObjectSummary(inserted)}をコピーしました。`);
    event.preventDefault?.();
    return;
	  }
  if (finishSelectedImageTransformForSheetGroup(job, image, event)) return;
	  if (job.mode === "rotate") {
	    setImageRotation(image, job.currentRotation);
	    if (image.shape) refreshShapeSource(image);
	    if (isLineShapeObject(image)) pruneDetachedLineShapeConnections(sheet, image);
	    else updateConnectedLinesForTargets(sheet, [image.id]);
	  } else if (job.mode === "shape-adjust") {
    image.shape = {
      ...(image.shape || {}),
      ...(customPathChanged ? { customPath: normalizeShapeCustomPath(job.currentCustomPath) } : {}),
      ...(adjustChanged ? { adjust: normalizeShapeAdjustments(job.shapeType, job.currentAdjust) } : {}),
    };
    refreshShapeSource(image);
  } else if ((job.mode === "move" && (job.imageIds || []).length > 1) || job.mode === "group-resize" || job.mode === "group-rotate") {
    const movedShapeIds = [];
    const movedLineIds = [];
    (job.imageIds || []).forEach((imageId) => {
      const target = sheet.images?.find((item) => item.id === imageId);
	      const box = job.currentBoxes?.[imageId];
	      if (!target || !box) return;
	      setImagePixelBox(sheet, target, box);
	      if (job.mode === "group-rotate") {
	        setImageRotation(target, job.currentRotations?.[imageId]);
	      }
      if (target.shape) refreshShapeSource(target);
      if (isLineShapeObject(target)) movedLineIds.push(target.id);
      else movedShapeIds.push(target.id);
    });
    updateConnectedLinesForTargets(sheet, movedShapeIds);
    pruneDetachedLineShapeConnectionsForIds(sheet, movedLineIds);
  } else {
    setImagePixelBox(sheet, image, job.currentBox);
    if (job.mode === "line-point" && image.shape) {
      image.shape.points = normalizeShapeLinePoints(job.currentLinePoints);
      setLineShapeConnections(image, updateLineShapeConnection(image.shape.connections, job.handle, job.currentLineConnection));
      pruneDetachedLineShapeConnections(sheet, image);
      refreshShapeSource(image);
    } else if (job.mode === "resize" && image.shape) {
      refreshShapeSource(image);
      if (isLineShapeObject(image)) pruneDetachedLineShapeConnections(sheet, image);
      else updateConnectedLinesForTargets(sheet, [image.id]);
    } else {
      if (isLineShapeObject(image)) pruneDetachedLineShapeConnections(sheet, image);
      else updateConnectedLinesForTargets(sheet, [image.id]);
    }
  }
  const actionLabel = job.mode === "rotate" || job.mode === "group-rotate" ? "回転" : job.mode === "line-point" ? "線の調整" : job.mode === "shape-adjust" ? "図形の調整" : job.mode === "group-resize" ? "グループ サイズ変更" : job.mode === "resize" ? "サイズ変更" : "移動";
  const statusAction = job.mode === "line-point" || job.mode === "shape-adjust" ? "調整" : actionLabel;
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore: job.sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${(job.imageIds || []).length > 1 ? `${job.imageIds.length} 個のオブジェクト` : image.name || "図形"} ${actionLabel}`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
	  setStatus(`${(job.imageIds || []).length > 1 ? `${job.imageIds.length} 個のオブジェクト` : image.name || "図形"}を${statusAction}しました。`);
	  event.preventDefault?.();
}



function finishSelectedImageTransformForSheetGroup(job, sourceImage, event) {
  const sourceSheet = activeSheet();
  const sourceIds = job.imageIds?.length ? job.imageIds : [job.imageId];
  const sourceImages = sourceIds.map((imageId) => sourceSheet.images?.find((item) => item.id === imageId)).filter(Boolean);
  const items = groupedShapeObjectTargets(sourceImages);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      const movedShapeIds = [];
      const movedLineIds = [];
      item.images.forEach((target, index) => {
        const source = sourceImages[index];
        if (!source) return;
        if (job.mode === "rotate") {
          setImageRotation(target, job.currentRotation);
          if (target.shape) refreshShapeSource(target);
          if (isLineShapeObject(target)) movedLineIds.push(target.id);
          else movedShapeIds.push(target.id);
          return;
        }
        if (job.mode === "shape-adjust") {
          target.shape = {
            ...(target.shape || {}),
            ...(job.currentCustomPath ? { customPath: normalizeShapeCustomPath(job.currentCustomPath) } : {}),
            ...(job.currentAdjust ? { adjust: normalizeShapeAdjustments(job.shapeType, job.currentAdjust) } : {}),
          };
          refreshShapeSource(target);
          return;
        }
        const sourceStartBox = job.startBoxes?.[source.id] || (source.id === job.imageId ? job.startBox : imageBoxToPixels(sourceSheet, source));
        const sourceCurrentBox = job.currentBoxes?.[source.id] || (source.id === job.imageId ? job.currentBox : null);
        const nextBox = transformedImageBoxForSheetTarget(item.sheet, target, sourceStartBox, sourceCurrentBox);
        if (nextBox) setImagePixelBox(item.sheet, target, nextBox);
        if (job.mode === "group-rotate") {
          setImageRotation(target, job.currentRotations?.[source.id]);
        }
        if (job.mode === "line-point" && target.shape) {
          target.shape.points = normalizeShapeLinePoints(job.currentLinePoints);
          setLineShapeConnections(target, updateLineShapeConnection(
            target.shape.connections,
            job.handle,
            lineConnectionForGroupedTarget(job.currentLineConnection, sourceSheet, item.sheet),
          ));
          pruneDetachedLineShapeConnections(item.sheet, target);
          refreshShapeSource(target);
        } else {
          if (target.shape) refreshShapeSource(target);
          if (isLineShapeObject(target)) movedLineIds.push(target.id);
          else movedShapeIds.push(target.id);
        }
      });
      updateConnectedLinesForTargets(item.sheet, movedShapeIds);
      pruneDetachedLineShapeConnectionsForIds(item.sheet, movedLineIds);
    });
    const actionLabel = job.mode === "rotate" || job.mode === "group-rotate" ? "回転" : job.mode === "line-point" ? "線の調整" : job.mode === "shape-adjust" ? "図形の調整" : job.mode === "group-resize" ? "グループ サイズ変更" : job.mode === "resize" ? "サイズ変更" : "移動";
    const statusAction = job.mode === "line-point" || job.mode === "shape-adjust" ? "調整" : actionLabel;
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート${(job.imageIds || []).length > 1 ? `${job.imageIds.length} 個のオブジェクト` : sourceImage.name || "図形"} ${actionLabel}`,
    });
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
    setStatus(`${items.length} 個のシートの${(job.imageIds || []).length > 1 ? `${job.imageIds.length} 個のオブジェクト` : sourceImage.name || "図形"}を${statusAction}しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクトを編集できませんでした：${error.message || error}`);
  }
  event?.preventDefault?.();
  return true;
}



function transformedImageBoxForSheetTarget(sheet, target, sourceStartBox, sourceCurrentBox) {
  if (!sourceStartBox || !sourceCurrentBox) return null;
  const targetStartBox = imageBoxToPixels(sheet, target);
  if (!targetStartBox) return null;
  return {
    left: targetStartBox.left + (sourceCurrentBox.left - sourceStartBox.left),
    top: targetStartBox.top + (sourceCurrentBox.top - sourceStartBox.top),
    width: Math.max(1, targetStartBox.width * (sourceCurrentBox.width / Math.max(1, sourceStartBox.width))),
    height: Math.max(1, targetStartBox.height * (sourceCurrentBox.height / Math.max(1, sourceStartBox.height))),
  };
}



function lineConnectionForGroupedTarget(connection, sourceSheet, targetSheet) {
  if (!connection?.imageId || !connection.site || targetSheet === sourceSheet) return connection || null;
  const source = sourceSheet.images?.find((image) => image.id === connection.imageId);
  const sourceIndex = sourceSheet.images?.indexOf(source);
  if (!source || sourceIndex < 0) return null;
  const target = correspondingGroupedShapeObject(targetSheet, sourceSheet, source, sourceIndex, new Set());
  return target ? { imageId: target.id, site: connection.site } : null;
}



function remapPastedLineShapeConnections(images, idMap, options = {}) {
  const map = idMap instanceof Map ? idMap : new Map();
  const preserveExternal = options.preserveExternal === true;
  (images || []).forEach((image) => {
    if (!isLineShapeObject(image)) return;
    const connections = normalizeLineShapeConnections(image.shape?.connections);
    if (!connections) return;
    const next = {};
    ["start", "end"].forEach((side) => {
      const connection = connections[side];
      if (!connection?.imageId || !connection.site) return;
      const mappedId = map.get(String(connection.imageId));
      if (mappedId) {
        next[side] = { imageId: mappedId, site: connection.site };
      } else if (preserveExternal) {
        next[side] = connection;
      }
    });
    const normalized = normalizeLineShapeConnections(next);
    if (normalized) image.shape.connections = normalized;
    else delete image.shape.connections;
  });
}



function setImageRotation(image, value) {
  if (!image) return;
  const rotation = normalizeRotationDegrees(value);
  if (image.shape) {
    image.shape = { ...image.shape, rotation };
  } else {
    image.rotation = rotation;
  }
}



function setImageFlip(image, axis, value) {
  if (!image) return;
  const key = axis === "v" ? "flipV" : "flipH";
  if (image.shape) {
    image.shape = { ...image.shape, [key]: Boolean(value) };
    refreshShapeSource(image);
  } else {
    image[key] = Boolean(value);
  }
}



function insertCopiedImagesFromMoveTransform(sheet, job) {
  const sourceIds = job.imageIds?.length ? job.imageIds : [job.imageId];
  const inserted = [];
  const groupMap = new Map();
  const copiedObjectIds = new Map();
  sourceIds.forEach((sourceId) => {
    const source = sheet.images?.find((item) => item.id === sourceId);
    if (!source) return;
    const nextBox = job.currentBoxes?.[sourceId] || (sourceId === job.imageId ? job.currentBox : null);
    if (!nextBox) return;
    const image = cloneImages([source])[0];
    const originalId = image.id;
    image.id = uniqueWorksheetObjectId("image");
    if (originalId) copiedObjectIds.set(String(originalId), image.id);
    image.name = copiedObjectName(source.name || source.shape?.label || "図形");
    if (image.shape) image.shape = { ...image.shape, label: image.name };
    if (image.groupId) {
      if (!groupMap.has(image.groupId)) groupMap.set(image.groupId, uniqueWorksheetObjectId("group"));
      image.groupId = groupMap.get(image.groupId);
    }
    setImagePixelBox(sheet, image, nextBox);
    if (image.shape) refreshShapeSource(image);
    sheet.images.push(image);
    inserted.push(image);
  });
  remapPastedLineShapeConnections(inserted, copiedObjectIds);
  if (inserted.length) {
    setSelectedImageIds(inserted.map((image) => image.id), inserted[0].id);
    state.selected = null;
    state.selectionRange = null;
    state.selectionRanges = [];
    state.selectionAnchor = null;
    state.selectionStart = null;
  }
  return inserted;
}



function imageTransformMoveDelta(job, pointer) {
  let dx = pointer.x - job.startPointer.x;
  let dy = pointer.y - job.startPointer.y;
  if (job.constrainMove) {
    if (Math.abs(dx) >= Math.abs(dy)) {
      dy = 0;
    } else {
      dx = 0;
    }
  }
  return { dx, dy };
}



function maybeSnapImageBoxTopLeftToGrid(box, enabled) {
  if (!enabled) return box;
  const sheet = activeSheet();
  const left = snapPixelToNearestGridLine(sheet, "col", box.left);
  const top = snapPixelToNearestGridLine(sheet, "row", box.top);
  return {
    ...box,
    left: Math.max(rowHeaderWidthForSheet(sheet), left),
    top: Math.max(COLUMN_HEADER_HEIGHT, top),
  };
}



function snapPixelToNearestGridLine(sheet, axis, pixel) {
  const index = sheetIndexFromPixel(sheet, axis, pixel);
  if (!index) return pixel;
  const before = coordinateToPixels(sheet, axis, index - 1);
  const after = coordinateToPixels(sheet, axis, index);
  return Math.abs(pixel - before) <= Math.abs(after - pixel) ? before : after;
}



function imageTransformBox(job, pointer) {
  const box = { ...job.startBox };
  const { dx, dy } = imageTransformMoveDelta(job, pointer);
  const minWidth = 12;
  const minHeight = 12;
  if (job.mode === "move") {
    const moved = {
      ...box,
      left: Math.max(rowHeaderWidthForSheet(activeSheet()), box.left + dx),
      top: Math.max(COLUMN_HEADER_HEIGHT, box.top + dy),
    };
    return maybeSnapImageBoxTopLeftToGrid(moved, Boolean(job.snapToGrid));
  }

  const right = box.left + box.width;
  const bottom = box.top + box.height;
  let left = box.left;
  let top = box.top;
  let nextRight = right;
  let nextBottom = bottom;
  if (job.handle.includes("w")) left = Math.min(right - minWidth, pointer.x);
  if (job.handle.includes("e")) nextRight = Math.max(left + minWidth, pointer.x);
  if (job.handle.includes("n")) top = Math.min(bottom - minHeight, pointer.y);
  if (job.handle.includes("s")) nextBottom = Math.max(top + minHeight, pointer.y);
  left = Math.max(rowHeaderWidthForSheet(activeSheet()), left);
  top = Math.max(COLUMN_HEADER_HEIGHT, top);
  const nextBox = {
    left,
    top,
    width: Math.max(minWidth, nextRight - left),
    height: Math.max(minHeight, nextBottom - top),
  };
  return constrainImageTransformAspect(nextBox, box, job.handle, Boolean(job.lockAspect));
}



function constrainImageTransformAspect(box, startBox, handle, lockAspect) {
  const hasHorizontal = handle.includes("w") || handle.includes("e");
  const hasVertical = handle.includes("n") || handle.includes("s");
  const aspect = Number(startBox.width) / Math.max(1, Number(startBox.height));
  if (!lockAspect || !hasHorizontal || !hasVertical || !Number.isFinite(aspect) || aspect <= 0) return box;

  const widthScale = box.width / Math.max(1, startBox.width);
  const heightScale = box.height / Math.max(1, startBox.height);
  let width = box.width;
  let height = box.height;
  if (Math.abs(widthScale - 1) >= Math.abs(heightScale - 1)) {
    height = Math.max(12, width / aspect);
  } else {
    width = Math.max(12, height * aspect);
  }

  const anchoredRight = handle.includes("w");
  const anchoredBottom = handle.includes("n");
  const right = anchoredRight ? startBox.left + startBox.width : box.left + width;
  const bottom = anchoredBottom ? startBox.top + startBox.height : box.top + height;
  return {
    left: anchoredRight ? right - width : box.left,
    top: anchoredBottom ? bottom - height : box.top,
    width,
    height,
  };
}



function shouldPreservePictureAspectOnResize(image, mode, handle) {
  return mode === "resize" && Boolean(image) && !image.shape && isCornerResizeHandle(handle);
}



function isCornerResizeHandle(handle) {
  return ["nw", "ne", "se", "sw"].includes(String(handle || ""));
}



function imageTransformMoveBoxes(job, pointer) {
  const { dx, dy } = imageTransformMoveDelta(job, pointer);
  const primaryBox = job.startBoxes?.[job.imageId];
  let adjustedDx = dx;
  let adjustedDy = dy;
  if (job.snapToGrid && primaryBox) {
    const primaryMoved = {
      ...primaryBox,
      left: Math.max(rowHeaderWidthForSheet(activeSheet()), primaryBox.left + dx),
      top: Math.max(COLUMN_HEADER_HEIGHT, primaryBox.top + dy),
    };
    const snapped = maybeSnapImageBoxTopLeftToGrid(primaryMoved, true);
    adjustedDx += snapped.left - primaryMoved.left;
    adjustedDy += snapped.top - primaryMoved.top;
  }
  return Object.fromEntries(
    Object.entries(job.startBoxes || {}).map(([imageId, box]) => [
      imageId,
      {
        ...box,
        left: Math.max(rowHeaderWidthForSheet(activeSheet()), box.left + adjustedDx),
        top: Math.max(COLUMN_HEADER_HEIGHT, box.top + adjustedDy),
      },
    ]),
  );
}



function imageTransformGroupResizeBoxes(job, pointer) {
  const startGroupBox = job.startGroupBox || job.startBox;
  const groupBox = imageTransformBox({ ...job, mode: "resize", startBox: startGroupBox }, pointer);
  const scaleX = groupBox.width / Math.max(1, startGroupBox.width);
  const scaleY = groupBox.height / Math.max(1, startGroupBox.height);
  const boxes = Object.fromEntries(
    Object.entries(job.startBoxes || {}).map(([imageId, box]) => [
      imageId,
      {
        left: groupBox.left + (box.left - startGroupBox.left) * scaleX,
        top: groupBox.top + (box.top - startGroupBox.top) * scaleY,
        width: Math.max(1, box.width * scaleX),
        height: Math.max(1, box.height * scaleY),
      },
    ]),
  );
  return { groupBox: imageTransformGroupVisualBox(job, boxes, job.startRotations) || groupBox, boxes };
}



function imageTransformGroupRotationBoxes(job, pointer) {
  const startGroupBox = job.startGroupBox || unionPixelBoxes(Object.values(job.startBoxes || {})) || job.startBox;
  const center = {
    x: startGroupBox.left + startGroupBox.width / 2,
    y: startGroupBox.top + startGroupBox.height / 2,
  };
  const startAngle = Math.atan2(job.startPointer.y - center.y, job.startPointer.x - center.x);
  const currentAngle = Math.atan2(pointer.y - center.y, pointer.x - center.x);
  let delta = ((currentAngle - startAngle) * 180) / Math.PI;
  if (job.lockAspect) delta = snapRotationDelta(delta);
  const boxes = Object.fromEntries(
    Object.entries(job.startBoxes || {}).map(([imageId, box]) => {
      const boxCenter = {
        x: box.left + box.width / 2,
        y: box.top + box.height / 2,
      };
      const rotated = rotatePointAround(boxCenter, center, delta);
      return [
        imageId,
        {
          ...box,
          left: rotated.x - box.width / 2,
          top: rotated.y - box.height / 2,
        },
      ];
    }),
  );
  const rotations = Object.fromEntries(
    Object.keys(job.startBoxes || {}).map((imageId) => [
      imageId,
      normalizeRotationDegrees((job.startRotations?.[imageId] || 0) + delta),
    ]),
  );
  return { boxes, rotations, groupBox: imageTransformGroupVisualBox(job, boxes, rotations) || startGroupBox };
}



function imageTransformGroupVisualBox(job, boxesById, rotationsById = null) {
  const sheet = activeSheet();
  const imageIds = new Set(Object.keys(boxesById || {}));
  const images = (sheet.images || []).filter((image) => imageIds.has(image.id));
  return unionImageSelectionBounds(sheet, images, boxesById, rotationsById) || unionPixelBoxes(Object.values(boxesById || {}));
}



function rotatePointAround(point, center, degrees) {
  const radians = (degrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const dx = point.x - center.x;
  const dy = point.y - center.y;
  return {
    x: center.x + dx * cos - dy * sin,
    y: center.y + dx * sin + dy * cos,
  };
}



function imageTransformRotation(job, pointer) {
  const box = job.startBox;
  const center = {
    x: box.left + box.width / 2,
    y: box.top + box.height / 2,
  };
  const startAngle = Math.atan2(job.startPointer.y - center.y, job.startPointer.x - center.x);
  const currentAngle = Math.atan2(pointer.y - center.y, pointer.x - center.x);
  let delta = ((currentAngle - startAngle) * 180) / Math.PI;
  if (job.lockAspect) delta = snapRotationDelta(delta);
  return normalizeRotationDegrees(job.startRotation + delta);
}



function snapRotationDelta(delta) {
  return Math.round((Number(delta) || 0) / 15) * 15;
}



function imageTransformLinePoint(job, pointer) {
  const image = selectedImageObjectById(job.imageId);
  const endpoints = lineEndpointPixelsFromBox(job.startBox, job.startLinePoints);
  const fixed = job.handle === "line-start" ? endpoints.end : endpoints.start;
  const rawMoved = {
    x: Math.max(rowHeaderWidthForSheet(activeSheet()), pointer.x),
    y: Math.max(COLUMN_HEADER_HEIGHT, pointer.y),
  };
  const snap = snapLineEndpointToShape(rawMoved, job.imageId);
  const previewTargetId = snap?.imageId || lineConnectionPreviewTarget(rawMoved, job.imageId);
  const moved = snap?.point || rawMoved;
  const result = job.handle === "line-start"
    ? lineBoxAndPointsFromEndpoints(moved, fixed, { shape: image?.shape || { type: job.shapeType } })
    : lineBoxAndPointsFromEndpoints(fixed, moved, { shape: image?.shape || { type: job.shapeType } });
  return { box: result.box, points: result.points, connection: snap, previewTargetId };
}



function snapLineEndpointToShape(point, lineImageId) {
  const sheet = activeSheet();
  const threshold = 14;
  let best = null;
  (sheet.images || []).forEach((image) => {
    if (!lineConnectionTargetObject(image, lineImageId)) return;
    shapeConnectionPointCandidates(sheet, image).forEach((candidate) => {
      const distance = Math.hypot(candidate.point.x - point.x, candidate.point.y - point.y);
      if (distance > threshold || (best && distance >= best.distance)) return;
      best = { imageId: image.id, site: candidate.site, point: candidate.point, distance };
    });
  });
  return best;
}



function lineConnectionPreviewTarget(point, lineImageId) {
  const sheet = activeSheet();
  const hitPadding = 6;
  const images = [...(sheet.images || [])].reverse();
  for (const image of images) {
    if (!lineConnectionTargetObject(image, lineImageId)) continue;
    const box = imageVisualBoxToPixels(image, imageBoxToPixels(sheet, image));
    if (!box) continue;
    if (
      point.x >= box.left - hitPadding &&
      point.x <= box.left + box.width + hitPadding &&
      point.y >= box.top - hitPadding &&
      point.y <= box.top + box.height + hitPadding
    ) {
      return image.id;
    }
  }
  return "";
}



function lineConnectionTargetObject(image, lineImageId = "") {
  return Boolean(image && !image.hidden && image.id !== lineImageId && !isLineShapeObject(image));
}



function updateLineConnectionPreviewOverlay(activeConnection, lineImageId, previewTargetId = "") {
  const grid = $sheetHost.find(".sheet-grid")[0];
  if (!grid || !state.model) return;
  const sheet = activeSheet();
  const activeImageId = activeConnection?.imageId || previewTargetId || "";
  if (!activeImageId) {
    removeLineConnectionPreviewOverlay();
    return;
  }
  let layer = grid.querySelector(".sheet-connection-preview-layer");
  if (!layer) {
    layer = document.createElement("div");
    layer.className = "sheet-connection-preview-layer";
    layer.setAttribute("aria-hidden", "true");
    grid.append(layer);
  }
  const visibleBounds = imageViewportBounds(sheet, visibleGridRange(sheet));
  const candidates = (sheet.images || [])
    .filter((image) => lineConnectionTargetObject(image, lineImageId))
    .filter((image) => {
      const box = imageBoxToPixels(sheet, image);
      return box && pixelBoxesIntersect(box, visibleBounds);
    })
    .filter((image) => image.id === activeImageId)
    .flatMap((image) =>
      shapeConnectionPointCandidates(sheet, image).map((candidate) => ({
        ...candidate,
        imageId: image.id,
      })),
    );
  layer.innerHTML = candidates
    .map((candidate) => {
      const active = candidate.imageId === activeImageId && connectionCandidateMatchesSite(candidate, activeConnection?.site);
      const className = `sheet-connection-point${active ? " is-active" : ""}`;
      return `<span class="${className}" style="left:${roundCssNumber(candidate.point.x)}px;top:${roundCssNumber(candidate.point.y)}px"></span>`;
    })
    .join("");
  state.lineConnectionPreview = activeConnection ? { ...activeConnection } : null;
}



function removeLineConnectionPreviewOverlay() {
  document.querySelector(".sheet-connection-preview-layer")?.remove();
  state.lineConnectionPreview = null;
}



function shapeConnectionSitesForImage(image) {
  const type = image?.shape?.type || DEFAULT_SHAPE_TYPE;
  const adjust = normalizeShapeAdjustments(type, image?.shape?.adjust) || {};
  switch (type) {
    case "rect":
    case "roundRect":
    case "snipRect":
    case "roundSnipRect":
    case "frame":
    case "textBox":
    case "flowProcess":
    case "flowTerminator":
    case "flowPredefinedProcess":
    case "flowInternalStorage":
    case "mathPlus":
    case "mathMinus":
    case "mathMultiply":
    case "mathDivide":
    case "mathEqual":
    case "mathNotEqual":
      return cardinalConnectionSites();
    case "oval":
    case "flowConnector":
    case "flowOr":
    case "flowSummingJunction":
      return cardinalConnectionSites();
    case "triangle":
      return sequentialConnectionSites([
        [0.5, 0, ["top"]],
        [0.25, 0.5],
        [0, 1, ["bottom-left"]],
        [0.5, 1, ["bottom"]],
        [1, 1, ["bottom-right"]],
        [0.75, 0.5],
      ]);
    case "rightTriangle":
      return sequentialConnectionSites([
        [0, 0, ["top-left"]],
        [0, 0.5, ["left"]],
        [0, 1, ["bottom-left"]],
        [0.5, 1, ["bottom"]],
        [1, 1, ["bottom-right"]],
        [0.5, 0.5],
      ]);
    case "diamond":
    case "flowDecision":
      return sequentialConnectionSites([
        [0.5, 0, ["top"]],
        [0, 0.5, ["left"]],
        [0.5, 1, ["bottom"]],
        [1, 0.5, ["right"]],
      ]);
    case "pentagon":
      return regularPolygonConnectionSites(5, -90, true);
    case "hexagon":
    case "flowPreparation":
      return polygonConnectionSites([
        [0.18, 0],
        [0.82, 0],
        [1, 0.5],
        [0.82, 1],
        [0.18, 1],
        [0, 0.5],
      ]);
    case "heptagon":
      return regularPolygonConnectionSites(7, -90, true);
    case "octagon":
      return polygonConnectionSites([
        [0.29, 0],
        [0.71, 0],
        [1, 0.3],
        [1, 0.7],
        [0.71, 1],
        [0.29, 1],
        [0, 0.7],
        [0, 0.3],
      ]);
    case "decagon":
      return regularPolygonConnectionSites(10, -90, true);
    case "dodecagon":
      return regularPolygonConnectionSites(12, -90, true);
    case "parallelogram":
      {
        const skew = clampNumber(adjust.skew, 0.04, 0.46, 0.22);
        return polygonConnectionSites([
          [skew, 0],
          [1, 0],
          [1 - skew, 1],
          [0, 1],
        ]);
      }
    case "trapezoid":
    case "flowManualOperation":
      {
        const inset = type === "trapezoid" ? clampNumber(adjust.inset, 0.04, 0.46, 0.24) : 0.18;
        return polygonConnectionSites([
          [inset, 0],
          [1 - inset, 0],
          [1, 1],
          [0, 1],
        ]);
      }
    case "flowData":
      {
        const skew = clampNumber(adjust.skew, 0.04, 0.44, 0.2);
        return polygonConnectionSites([
          [skew, 0],
          [1, 0],
          [1 - skew, 1],
          [0, 1],
        ]);
      }
    case "flowManualInput":
      {
        const skew = clampNumber(adjust.skew, 0.04, 0.44, 0.12);
        return polygonConnectionSites([
          [0, skew],
          [1, 0],
          [1, 1],
          [0, 1],
        ]);
      }
    case "cross":
      return cardinalConnectionSites();
    case "can":
      return sequentialConnectionSites([
        [0.5, 0, ["top"]],
        [0, 0.5, ["left"]],
        [0.5, 1, ["bottom"]],
        [1, 0.5, ["right"]],
      ]);
    case "cube":
      return polygonConnectionSites([
        [0, 0.22],
        [0.22, 0],
        [1, 0],
        [1, 0.77],
        [0.78, 1],
        [0, 1],
      ]);
    case "bevel":
    case "donut":
    case "noSymbol":
    case "blockArc":
    case "foldedCorner":
    case "smileyFace":
    case "heart":
    case "sun":
    case "moon":
    case "cloud":
    case "pie":
    case "chord":
    case "teardrop":
    case "plaque":
    case "wave":
      return cardinalConnectionSites();
    case "lightningBolt":
      return polygonConnectionSites([
        [0.55, 0],
        [0.18, 0.57],
        [0.43, 0.57],
        [0.34, 1],
        [0.82, 0.37],
        [0.55, 0.37],
      ]);
    case "corner":
    case "halfFrame":
      return polygonConnectionSites([
        [0, 0],
        [1, 0],
        [1, 0.3],
        [0.18, 0.3],
        [0.18, 1],
        [0, 1],
      ]);
    case "diagStripe":
      return polygonConnectionSites([
        [0, 1],
        [0.6, 0],
        [1, 0],
        [0.4, 1],
      ]);
    case "leftBracket":
    case "rightBracket":
    case "bracketPair":
    case "leftBrace":
    case "rightBrace":
    case "bracePair":
    case "arc":
      return cardinalConnectionSites();
    case "rightBlockArrow":
      return blockArrowConnectionSites("right", adjust);
    case "leftBlockArrow":
      return blockArrowConnectionSites("left", adjust);
    case "upBlockArrow":
      return blockArrowConnectionSites("up", adjust);
    case "downBlockArrow":
      return blockArrowConnectionSites("down", adjust);
    case "leftRightBlockArrow":
      return sequentialConnectionSites([
        [1, 0.5, ["right"]],
        [0.72, 1],
        [0.28, 1],
        [0, 0.5, ["left"]],
        [0.28, 0],
        [0.72, 0],
      ]);
    case "leftRightUpArrow":
      return sequentialConnectionSites([
        [0.5, 0, ["top"]],
        [1, 0.5, ["right"]],
        [0.72, 1],
        [0.28, 1],
        [0, 0.5, ["left"]],
      ]);
    case "quadArrow":
      return cardinalConnectionSites();
    case "notchedRightArrow":
      return sequentialConnectionSites([
        [clampNumber(adjust.head, 0.22, 0.82, 0.6), 0, ["top"]],
        [0, 0.5, ["left"]],
        [clampNumber(adjust.head, 0.22, 0.82, 0.6), 1, ["bottom"]],
        [1, 0.5, ["right"]],
      ]);
    case "bentUpArrow":
    case "uturnArrow":
    case "curvedRightArrow":
    case "leftUpArrow":
    case "stripedRightArrow":
    case "homePlate":
    case "chevron":
    case "leftRightCircularArrow":
      return cardinalConnectionSites();
    case "flowDocument":
    case "flowMultiDocument":
      return sequentialConnectionSites([
        [0.5, 0, ["top"]],
        [0, 0.5, ["left"]],
        [0.5, 0.88, ["bottom"]],
        [1, 0.5, ["right"]],
      ]);
    case "flowStoredData":
    case "flowDelay":
    case "flowDisplay":
      return cardinalConnectionSites();
    case "flowOffpageConnector":
      return polygonConnectionSites([
        [0, 0],
        [1, 0],
        [1, 0.68],
        [0.5, 1],
        [0, 0.68],
      ]);
    case "star4":
      return cardinalConnectionSites();
    case "star5":
    case "star6":
    case "star7":
    case "star8":
    case "star10":
    case "star12":
    case "star16":
    case "star24":
    case "star32":
      return regularPolygonConnectionSites(starPointCountForType(type), -90, true);
    case "seal":
      return compassConnectionSites(svgPointsStringToNormalizedPoints(explosionSealPoints()));
    case "explosion":
      return compassConnectionSites(svgPointsStringToNormalizedPoints(explosionBurstPoints()));
    case "ribbon":
      return polygonConnectionSites([
        [0, 0],
        [1, 0],
        [1, 0.75],
        [0.66, 0.75],
        [0.5, 1],
        [0.34, 0.75],
        [0, 0.75],
      ]);
    case "calloutRect":
    case "calloutRoundRect":
    case "calloutOval":
    case "calloutCloud":
      return calloutConnectionSites(type, adjust);
    case CUSTOM_SHAPE_TYPE:
      return customShapeConnectionSites(image.shape?.customPath);
    default:
      return cardinalConnectionSites();
  }
}



function connectionSite(x, y, aliases = []) {
  return {
    x: Number.isFinite(Number(x)) ? Number(x) : 0,
    y: Number.isFinite(Number(y)) ? Number(y) : 0,
    aliases: aliases.filter(Boolean),
  };
}



function sequentialConnectionSites(points) {
  return points.map(([x, y, aliases = []], index) => ({
    ...connectionSite(x, y, aliases),
    site: `connection-${index}`,
  }));
}



function cardinalConnectionSites() {
  return sequentialConnectionSites([
    [0.5, 0, ["top"]],
    [0, 0.5, ["left"]],
    [0.5, 1, ["bottom"]],
    [1, 0.5, ["right"]],
  ]);
}



function polygonConnectionSites(points) {
  return sequentialConnectionSites(points.map(([x, y], index) => [x, y, polygonPointAliases(x, y, index)]));
}



function polygonPointAliases(x, y, index) {
  const aliases = [];
  if (index === 0 && Math.abs(y) < 0.001) aliases.push("top");
  if (Math.abs(x) < 0.001) aliases.push("left");
  if (Math.abs(y - 1) < 0.001) aliases.push("bottom");
  if (Math.abs(x - 1) < 0.001) aliases.push("right");
  return aliases;
}



function regularPolygonConnectionSites(count, startDegrees = -90, reverse = false) {
  const direction = reverse ? -1 : 1;
  const raw = Array.from({ length: Math.max(3, count) }, (_value, index) => {
    const angle = ((startDegrees + direction * (360 / count) * index) * Math.PI) / 180;
    return [50 + Math.cos(angle) * 31, 30 + Math.sin(angle) * 31];
  });
  return polygonConnectionSites(fitViewBoxPoints(raw).map(([x, y]) => [x / 100, y / 60]));
}



function fitViewBoxPoints(points) {
  const xs = points.map(([x]) => x);
  const ys = points.map(([, y]) => y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  const width = Math.max(0.0001, maxX - minX);
  const height = Math.max(0.0001, maxY - minY);
  return points.map(([x, y]) => [((x - minX) / width) * 100, ((y - minY) / height) * 60]);
}



function blockArrowConnectionSites(direction, adjust = {}) {
  const head = clampNumber(adjust.head, 0.22, 0.82, direction === "left" ? 0.4 : direction === "down" ? 0.48 : 0.52);
  if (direction === "right") {
    return sequentialConnectionSites([
      [head, 0, ["top"]],
      [0, 0.5, ["left"]],
      [head, 1, ["bottom"]],
      [1, 0.5, ["right"]],
    ]);
  }
  if (direction === "left") {
    return sequentialConnectionSites([
      [head, 0, ["top"]],
      [0, 0.5, ["left"]],
      [head, 1, ["bottom"]],
      [1, 0.5, ["right"]],
    ]);
  }
  if (direction === "up") {
    return sequentialConnectionSites([
      [0.5, 0, ["top"]],
      [0, head, ["left"]],
      [0.5, 1, ["bottom"]],
      [1, head, ["right"]],
    ]);
  }
  return sequentialConnectionSites([
    [0.5, 0, ["top"]],
    [0, head, ["left"]],
    [0.5, 1, ["bottom"]],
    [1, head, ["right"]],
  ]);
}



function calloutConnectionSites(type, adjust = {}) {
  const normalized = normalizeShapeAdjustments(type, adjust) || {};
  return [
    ...cardinalConnectionSites(),
    {
      site: "connection-4",
      x: clampNumber(normalized.tailX, -2, 3, defaultCalloutTailX(type)),
      y: clampNumber(normalized.tailY, -2, 3, defaultCalloutTailY(type)),
      aliases: ["tail"],
    },
  ];
}



function svgPointsStringToNormalizedPoints(pointsString) {
  return String(pointsString || "")
    .trim()
    .split(/\s+/)
    .map((pair) => pair.split(",").map(Number))
    .filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y))
    .map(([x, y]) => [x / 100, y / 60]);
}



function compassConnectionSites(points) {
  if (!points?.length) return cardinalConnectionSites();
  const center = { x: 0.5, y: 0.5 };
  const pick = (score) =>
    points.reduce((best, point) => {
      if (!best) return point;
      const bestScore = score(best);
      const pointScore = score(point);
      if (pointScore < bestScore) return point;
      if (pointScore === bestScore && Math.hypot(point[0] - center.x, point[1] - center.y) < Math.hypot(best[0] - center.x, best[1] - center.y)) return point;
      return best;
    }, null);
  const top = pick((point) => point[1]);
  const left = pick((point) => point[0]);
  const bottom = pick((point) => -point[1]);
  const right = pick((point) => -point[0]);
  return sequentialConnectionSites([
    [top[0], top[1], ["top"]],
    [left[0], left[1], ["left"]],
    [bottom[0], bottom[1], ["bottom"]],
    [right[0], right[1], ["right"]],
  ]);
}



function customShapeConnectionSites(customPath) {
  const path = normalizeShapeCustomPath(customPath);
  const points = (path?.commands || [])
    .flatMap((command) => command.points || [])
    .map((point) => [point.x, point.y]);
  return points.length ? polygonConnectionSites(points) : cardinalConnectionSites();
}



function shapeConnectionPointCandidates(sheet, image, options = {}) {
  const box = options.box || imageBoxToPixels(sheet, image);
  if (!box) return [];
  const raw = shapeConnectionSitesForImage(image);
  const center = { x: box.left + box.width / 2, y: box.top + box.height / 2 };
  const rotation = options.rotation === undefined ? imageRotationDegrees(image) : normalizeRotationDegrees(options.rotation);
  return raw.map((site, index) => {
    const x = imageFlipH(image) ? 1 - site.x : site.x;
    const y = imageFlipV(image) ? 1 - site.y : site.y;
    const point = {
      x: box.left + box.width * x,
      y: box.top + box.height * y,
    };
    return {
      site: site.site || `connection-${index}`,
      aliases: site.aliases || [],
      point: rotation ? rotatePointAround(point, center, rotation) : point,
    };
  });
}



function shapeConnectionPointForSite(sheet, image, site, options = {}) {
  const candidates = shapeConnectionPointCandidates(sheet, image, options);
  const direct = candidates.find((candidate) => connectionCandidateMatchesSite(candidate, site))?.point;
  if (direct) return direct;
  const indexed = String(site || "").match(/^connection-(\d+)$/);
  if (indexed && candidates.length) {
    const index = Number(indexed[1]);
    return candidates[index % candidates.length]?.point || null;
  }
  return null;
}



function connectionCandidateMatchesSite(candidate, site) {
  const value = String(site || "");
  if (!value) return false;
  return candidate?.site === value || (candidate?.aliases || []).includes(value);
}



function updateLineShapeConnection(connections, handle, connection) {
  const side = handle === "line-start" ? "start" : "end";
  const next = { ...(connections || {}) };
  if (connection?.imageId && connection.site) {
    next[side] = { imageId: connection.imageId, site: connection.site };
  } else {
    delete next[side];
  }
  return next.start || next.end ? next : null;
}



function setLineShapeConnections(image, connections) {
  if (!image?.shape) return false;
  const normalized = normalizeLineShapeConnections(connections);
  if (normalized) {
    image.shape.connections = normalized;
  } else {
    delete image.shape.connections;
  }
  return true;
}



function pruneDetachedLineShapeConnections(sheet, image, tolerance = LINE_CONNECTION_RETENTION_TOLERANCE_PX) {
  if (!sheet || !isLineShapeObject(image)) return false;
  const connections = normalizeLineShapeConnections(image.shape?.connections);
  if (!connections) {
    if (image?.shape && "connections" in image.shape) delete image.shape.connections;
    return false;
  }
  const box = imageBoxToPixels(sheet, image);
  if (!box) return false;
  const endpoints = lineEndpointPixelsFromBox(box, normalizeShapeLinePoints(image.shape?.points));
  const next = {};
  let changed = false;
  ["start", "end"].forEach((side) => {
    const connection = connections[side];
    if (!connection?.imageId || !connection.site) return;
    const target = sheet.images?.find((item) => item.id === connection.imageId);
    const point = target ? shapeConnectionPointForSite(sheet, target, connection.site) : null;
    const endpoint = endpoints[side];
    const keep = point && endpoint && Math.hypot(point.x - endpoint.x, point.y - endpoint.y) <= tolerance;
    if (keep) {
      next[side] = connection;
    } else {
      changed = true;
    }
  });
  if (!changed) return false;
  setLineShapeConnections(image, next);
  return true;
}



function pruneDetachedLineShapeConnectionsForIds(sheet, imageIds) {
  let changed = 0;
  (imageIds || []).forEach((imageId) => {
    const image = sheet?.images?.find((item) => item.id === imageId);
    if (pruneDetachedLineShapeConnections(sheet, image)) changed += 1;
  });
  return changed;
}



function updateConnectedLinesForTargets(sheet, targetIds) {
  const targetSet = new Set((targetIds || []).filter(Boolean));
  if (!sheet?.images?.length || !targetSet.size) return 0;
  let updated = 0;
  sheet.images.forEach((image) => {
    if (!isLineShapeObject(image)) return;
    const connections = normalizeLineShapeConnections(image.shape?.connections);
    if (!connections) return;
    const touchesTarget = ["start", "end"].some((side) => targetSet.has(connections[side]?.imageId));
    if (!touchesTarget) return;
    const box = imageBoxToPixels(sheet, image);
    if (!box) return;
    const endpoints = lineEndpointPixelsFromBox(box, normalizeShapeLinePoints(image.shape?.points));
    const next = {
      start: { ...endpoints.start },
      end: { ...endpoints.end },
    };
    let changed = false;
    ["start", "end"].forEach((side) => {
      const connection = connections[side];
      if (!targetSet.has(connection?.imageId)) return;
      const target = sheet.images?.find((item) => item.id === connection.imageId);
      const point = target ? shapeConnectionPointForSite(sheet, target, connection.site) : null;
      if (!point) return;
      next[side] = point;
      changed = true;
    });
    if (!changed) return;
    const result = lineBoxAndPointsFromEndpoints(next.start, next.end, { shape: image.shape });
    setImagePixelBox(sheet, image, result.box);
    image.shape = {
      ...(image.shape || {}),
      points: normalizeShapeLinePoints(result.points),
      connections,
    };
    refreshShapeSource(image);
    updated += 1;
  });
  return updated;
}



function previewConnectedLinesForTransform(job) {
  if (!job || job.mode === "line-point") return;
  const sheet = activeSheet();
  const targetIds = transformConnectionTargetIds(job, sheet);
  if (!targetIds.length) return;
  const targetSet = new Set(targetIds);
  const boxOverrides = transformConnectionBoxOverrides(job);
  const rotationOverrides = transformConnectionRotationOverrides(job);
  (sheet.images || []).forEach((line) => {
    if (!isLineShapeObject(line)) return;
    const connections = normalizeLineShapeConnections(line.shape?.connections);
    if (!connections) return;
    if (!["start", "end"].some((side) => targetSet.has(connections[side]?.imageId))) return;
    const lineBox = imageBoxToPixels(sheet, line);
    if (!lineBox) return;
    const endpoints = lineEndpointPixelsFromBox(lineBox, normalizeShapeLinePoints(line.shape?.points));
    const next = {
      start: { ...endpoints.start },
      end: { ...endpoints.end },
    };
    let changed = false;
    ["start", "end"].forEach((side) => {
      const connection = connections[side];
      if (!targetSet.has(connection?.imageId)) return;
      const target = sheet.images?.find((item) => item.id === connection.imageId);
      const point = target
        ? shapeConnectionPointForSite(sheet, target, connection.site, {
            box: boxOverrides[connection.imageId],
            rotation: rotationOverrides[connection.imageId],
          })
        : null;
      if (!point) return;
      next[side] = point;
      changed = true;
    });
    if (!changed) return;
    const result = lineBoxAndPointsFromEndpoints(next.start, next.end, { shape: line.shape });
    previewSelectedImageBox(line.id, result.box, { points: result.points });
  });
}



function transformConnectionTargetIds(job, sheet) {
  const ids = job.imageIds?.length ? job.imageIds : [job.imageId];
  const boxOverrides = transformConnectionBoxOverrides(job);
  return ids.filter((imageId) => {
    const image = sheet.images?.find((item) => item.id === imageId);
    return lineConnectionTargetObject(image) && Boolean(boxOverrides[imageId] || imageId === job.imageId);
  });
}



function transformConnectionBoxOverrides(job) {
  if (!job) return {};
  const boxes = job.currentBoxes && Object.keys(job.currentBoxes).length ? { ...job.currentBoxes } : {};
  if (job.imageId && job.currentBox) boxes[job.imageId] = job.currentBox;
  return boxes;
}



function transformConnectionRotationOverrides(job) {
  if (!job) return {};
  const rotations = job.currentRotations && Object.keys(job.currentRotations).length ? { ...job.currentRotations } : {};
  if (job.imageId && job.currentRotation !== undefined) rotations[job.imageId] = job.currentRotation;
  return rotations;
}



function imageTransformShapeAdjustment(job, pointer) {
  const box = job.startBox;
  const isCallout = isCalloutShapeType(job.shapeType);
  const point = {
    x: isCallout
      ? clampNumber((pointer.x - box.left) / Math.max(1, box.width), -2, 3, 0.5)
      : clamp01((pointer.x - box.left) / Math.max(1, box.width)),
    y: isCallout
      ? clampNumber((pointer.y - box.top) / Math.max(1, box.height), -2, 3, 0.5)
      : clamp01((pointer.y - box.top) / Math.max(1, box.height)),
  };
  return shapeAdjustmentFromPointer(job.shapeType, job.startAdjust, job.handle, point, {
    box,
    points: job.startLinePoints,
  });
}



function imageTransformCustomPathPoint(job, pointer) {
  const path = normalizeShapeCustomPath(job.startCustomPath);
  if (!path) return null;
  const match = String(job.handle || "").match(/^custom-point:(\d+):(\d+)$/);
  if (!match) return path;
  const commandIndex = Number(match[1]);
  const pointIndex = Number(match[2]);
  const point = {
    x: clamp01((pointer.x - job.startBox.left) / Math.max(1, job.startBox.width)),
    y: clamp01((pointer.y - job.startBox.top) / Math.max(1, job.startBox.height)),
  };
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return path;
  const commands = path.commands.map((command, index) => {
    if (index !== commandIndex) return command;
    return {
      ...command,
      points: command.points.map((item, subIndex) => (subIndex === pointIndex ? point : item)),
    };
  });
  return normalizeShapeCustomPath({ commands }) || path;
}



function shapeAdjustmentFromPointer(type, startAdjust, handle, point, context = {}) {
  const next = { ...(normalizeShapeAdjustments(type, startAdjust) || {}) };
  if (!Object.keys(next).length) return null;
  switch (handle) {
    case "bend":
      if (type === "elbowConnector") {
        const horizontalMiddle = elbowConnectorUsesHorizontalMiddle({ points: context.points, adjust: next }, context.box);
        next.bend = clampElbowConnectorBendForPoints(context.points, horizontalMiddle ? point.y : point.x, horizontalMiddle);
      } else {
        next.bend = point.x;
      }
      break;
    case "curve":
      next.curve = point.y;
      break;
    case "corner":
      next.corner = type === "roundRect"
        ? point.x / roundRectCornerScale(context.box?.width, context.box?.height).x
        : point.x;
      break;
    case "cut":
      next.cut = 1 - point.x;
      break;
    case "skew":
      next.skew = type === "flowManualInput" ? point.y : point.x;
      break;
    case "inset":
      next.inset = point.x;
      break;
    case "thickness":
      next.thickness = ["upBlockArrow", "downBlockArrow"].includes(type)
        ? 1 - point.x * 2
        : 1 - point.y * 2;
      break;
    case "head":
      next.head = ["upBlockArrow", "downBlockArrow"].includes(type) ? point.y : point.x;
      break;
    case "depth":
      next.depth = point.y;
      break;
    case "inner":
      next.inner = (point.x - 0.5) / 0.34;
      break;
    case "tail":
      next.tailX = point.x;
      next.tailY = point.y;
      break;
    case "tailHeight":
      next.tailHeight = Math.max(0.08, next.tailY - point.y);
      break;
    default:
      break;
  }
  return normalizeShapeAdjustments(type, next);
}



function lineEndpointPixelsFromBox(box, points) {
  const normalized = normalizeShapeLinePoints(points);
  return {
    start: {
      x: box.left + box.width * normalized.x1,
      y: box.top + box.height * normalized.y1,
    },
    end: {
      x: box.left + box.width * normalized.x2,
      y: box.top + box.height * normalized.y2,
    },
  };
}



function lineShapePointerHitsStroke(sheet, image, event) {
  const point = sheetPixelFromClientPosition(event.clientX, event.clientY, { clampToGrid: false });
  if (!point) return false;
  const segments = lineShapeSegmentsInSheetPixels(sheet, image);
  if (!segments.length) return false;
  const width = normalizeShapeOutlineWidth(image?.shape?.outlineWidth);
  const tolerance = Math.max(7, width * 2.2 + 5);
  return segments.some(([start, end]) => pointToSegmentDistance(point, start, end) <= tolerance);
}



function lineShapeSegmentsInSheetPixels(sheet, image, boxOverride = null) {
  const box = boxOverride || imageBoxToPixels(sheet, image);
  if (!box) return [];
  const points = lineShapeRoutePointsLocal(image, box);
  if (points.length < 2) return [];
  const center = { x: box.left + box.width / 2, y: box.top + box.height / 2 };
  const rotation = imageRotationDegrees(image);
  const sheetPoints = points.map((point) => {
    const sheetPoint = { x: box.left + point.x, y: box.top + point.y };
    return rotation ? rotatePointAround(sheetPoint, center, rotation) : sheetPoint;
  });
  return sheetPoints.slice(1).map((point, index) => [sheetPoints[index], point]);
}



function lineShapeRoutePointsLocal(image, box) {
  const shape = image?.shape || {};
  const points = normalizeShapeLinePoints(shape.points);
  const p1 = linePointToPixelPoint(points.x1, points.y1, box.width, box.height);
  const p2 = linePointToPixelPoint(points.x2, points.y2, box.width, box.height);
  switch (shape.type) {
    case "elbowConnector":
      return elbowConnectorRoutePointsPx(p1, p2, shape.adjust, box.width, box.height);
    case "curvedConnector":
      {
        const curveOffset = roundSvgNumber((0.5 - (shape.adjust?.curve ?? 0.16)) * box.height * 0.86);
        const c1 = { x: roundSvgNumber(p1.x + (p2.x - p1.x) * 0.38), y: roundSvgNumber(p1.y + curveOffset) };
        const c2 = { x: roundSvgNumber(p1.x + (p2.x - p1.x) * 0.62), y: roundSvgNumber(p2.y + curveOffset) };
        return sampleCubicBezierPoints(p1, c1, c2, p2, 18);
      }
    case "freeform":
      {
        const point = (x, y) => ({ x: roundSvgNumber(box.width * x), y: roundSvgNumber(box.height * y) });
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
        return [
          ...sampleCubicBezierPoints(controls[0], controls[1], controls[2], controls[3], 10),
          ...sampleCubicBezierPoints(controls[3], controls[4], controls[5], controls[6], 10).slice(1),
          ...sampleCubicBezierPoints(controls[6], controls[6], controls[6], controls[7], 4).slice(1),
        ];
      }
    default:
      return [p1, p2];
  }
}



function sampleCubicBezierPoints(p0, p1, p2, p3, steps = 12) {
  const count = Math.max(2, Math.round(Number(steps) || 12));
  const points = [];
  for (let i = 0; i <= count; i += 1) {
    const t = i / count;
    const mt = 1 - t;
    points.push({
      x: mt ** 3 * p0.x + 3 * mt ** 2 * t * p1.x + 3 * mt * t ** 2 * p2.x + t ** 3 * p3.x,
      y: mt ** 3 * p0.y + 3 * mt ** 2 * t * p1.y + 3 * mt * t ** 2 * p2.y + t ** 3 * p3.y,
    });
  }
  return points;
}



function pointToSegmentDistance(point, start, end) {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const lengthSquared = dx * dx + dy * dy;
  if (!lengthSquared) return Math.hypot(point.x - start.x, point.y - start.y);
  const t = clampNumber(((point.x - start.x) * dx + (point.y - start.y) * dy) / lengthSquared, 0, 1, 0);
  return Math.hypot(point.x - (start.x + dx * t), point.y - (start.y + dy * t));
}



function previewSelectedImageBox(imageId, box, options = {}) {
  const image = selectedImageObjectById(imageId) || activeSheet().images?.find((item) => item.id === imageId);
	  const previewImage = image?.id === imageId
	    ? {
	        ...image,
	        ext: box ? { width: Math.max(1, box.width), height: Math.max(1, box.height) } : image.ext,
	        ...(options.rotation !== undefined && !image.shape ? { rotation: options.rotation } : {}),
	        shape: image.shape
	          ? {
	              ...image.shape,
              ...(options.rotation !== undefined ? { rotation: options.rotation } : {}),
              ...(options.points ? { points: normalizeShapeLinePoints(options.points) } : {}),
              ...(options.adjust ? { adjust: normalizeShapeAdjustments(image.shape.type, options.adjust) } : {}),
              ...(options.customPath ? { customPath: normalizeShapeCustomPath(options.customPath) } : {}),
            }
          : image.shape,
      }
    : null;
  const styleText = imageBoxStyle(previewImage || {}, box);
  const frameStyleText = imageSelectionFrameStyle(previewImage || {}, box);
  const style = Object.fromEntries(styleText.split(";").filter(Boolean).map((item) => {
    const index = item.indexOf(":");
    return [item.slice(0, index), item.slice(index + 1)];
  }));
  const frameStyle = Object.fromEntries(frameStyleText.split(";").filter(Boolean).map((item) => {
    const index = item.indexOf(":");
    return [item.slice(0, index), item.slice(index + 1)];
  }));
  const frame = $sheetHost.find(`.sheet-image-selection-frame[data-image-id="${cssEscape(imageId)}"]:not([data-image-group-id])`)[0];
  if (frame) {
    Object.assign(frame.style, frameStyle);
    if (options.points) updateLineHandlePositions(frame, options.points);
    if (previewImage?.shape && options.adjust) updateShapeAdjustmentHandlePositions(frame, previewImage.shape);
    if (previewImage?.shape && options.customPath) updateShapeAdjustmentHandlePositions(frame, previewImage.shape);
  }
  const element = primaryImagePositionElement(imageId);
  if (element) {
    Object.assign(element.style, style);
    const previewSizeChanged =
      previewImage?.shape &&
      box &&
      (Math.round(Number(box.width) || 0) !== Math.round(Number(image?.ext?.width) || 0) ||
        Math.round(Number(box.height) || 0) !== Math.round(Number(image?.ext?.height) || 0));
    if (previewImage?.shape && (options.points || options.adjust || options.customPath || previewSizeChanged)) {
      refreshShapeSource(previewImage);
      setImageElementRenderSource(element, previewImage.src);
    }
  }
}



function previewSelectedImageGroupFrame(groupId, boxesById, rotationsById = null) {
  const sheet = activeSheet();
  const imageIds = new Set(Object.keys(boxesById || {}));
  const images = (sheet.images || []).filter((image) => image.groupId === groupId && imageIds.has(image.id));
  const box = images.length
    ? unionImageSelectionBounds(sheet, images, boxesById, rotationsById)
    : unionPixelBoxes(Object.values(boxesById || {}));
  if (!box) return;
  const frame = $sheetHost.find(`.sheet-image-selection-frame[data-image-group-id="${cssEscape(groupId)}"]`)[0];
  if (!frame) return;
  Object.assign(frame.style, {
    left: `${roundCssNumber(box.left)}px`,
    top: `${roundCssNumber(box.top)}px`,
    width: `${roundCssNumber(box.width)}px`,
    height: `${roundCssNumber(box.height)}px`,
  });
}



function unionPixelBoxes(boxes) {
  const valid = (boxes || []).filter(Boolean);
  if (!valid.length) return null;
  const left = Math.min(...valid.map((box) => box.left));
  const top = Math.min(...valid.map((box) => box.top));
  const right = Math.max(...valid.map((box) => box.left + box.width));
  const bottom = Math.max(...valid.map((box) => box.top + box.height));
  return { left, top, width: Math.max(1, right - left), height: Math.max(1, bottom - top) };
}



function updateLineHandlePositions(frame, points) {
  const normalized = normalizeShapeLinePoints(points);
  const start = frame.querySelector('[data-image-resize="line-start"]');
  const end = frame.querySelector('[data-image-resize="line-end"]');
  if (start) {
    start.style.left = `${roundCssNumber(normalized.x1 * 100)}%`;
    start.style.top = `${roundCssNumber(normalized.y1 * 100)}%`;
  }
  if (end) {
    end.style.left = `${roundCssNumber(normalized.x2 * 100)}%`;
    end.style.top = `${roundCssNumber(normalized.y2 * 100)}%`;
  }
}



function updateShapeAdjustmentHandlePositions(frame, shape) {
  if (shape?.type === CUSTOM_SHAPE_TYPE) {
    shapeCustomPathHandles(shape.customPath).forEach((handle) => {
      const element = frame.querySelector(`[data-shape-adjust="${cssEscape(handle.key)}"]`);
      if (!element) return;
      element.style.left = `${roundCssNumber(handle.x * 100)}%`;
      element.style.top = `${roundCssNumber(handle.y * 100)}%`;
    });
    return;
  }
  shapeAdjustmentHandlesForFrame(shape, shapeAdjustmentBoxFromFrame(frame)).forEach((handle) => {
    const element = frame.querySelector(`[data-shape-adjust="${cssEscape(handle.key)}"]`);
    if (!element) return;
    element.style.left = `${roundCssNumber(handle.x * 100)}%`;
    element.style.top = `${roundCssNumber(handle.y * 100)}%`;
  });
}



function sameRotation(a, b) {
  return Math.abs(normalizeRotationDegrees(a) - normalizeRotationDegrees(b)) < 0.1;
}



function linePointsEqual(a, b) {
  const left = normalizeShapeLinePoints(a);
  const right = normalizeShapeLinePoints(b);
  return ["x1", "y1", "x2", "y2"].every((key) => Math.abs(left[key] - right[key]) < 0.001);
}



function shapeAdjustmentsEqual(type, a, b) {
  const left = normalizeShapeAdjustments(type, a);
  const right = normalizeShapeAdjustments(type, b);
  if (!left && !right) return true;
  if (!left || !right) return false;
  const keys = new Set([...Object.keys(left), ...Object.keys(right)]);
  return [...keys].every((key) => Math.abs(Number(left[key]) - Number(right[key])) < 0.001);
}



function shapeCustomPathsEqual(a, b) {
  const left = normalizeShapeCustomPath(a);
  const right = normalizeShapeCustomPath(b);
  if (!left && !right) return true;
  if (!left || !right) return false;
  if (left.commands.length !== right.commands.length) return false;
  return left.commands.every((command, index) => {
    const other = right.commands[index];
    if (command.cmd !== other.cmd || command.points.length !== other.points.length) return false;
    return command.points.every((point, pointIndex) => Math.abs(point.x - other.points[pointIndex].x) < 0.001 && Math.abs(point.y - other.points[pointIndex].y) < 0.001);
  });
}



function primaryImagePositionElement(imageId) {
  const escaped = cssEscape(imageId);
  return (
    $sheetHost.find(`.sheet-image-link[data-image-id="${escaped}"]`)[0] ||
    $sheetHost.find(`.sheet-image[data-image-id="${escaped}"]`)[0] ||
    null
  );
}



function boxesEqual(a, b) {
  return Boolean(
    a &&
      b &&
      Math.round(a.left) === Math.round(b.left) &&
      Math.round(a.top) === Math.round(b.top) &&
      Math.round(a.width) === Math.round(b.width) &&
      Math.round(a.height) === Math.round(b.height),
  );
}



function setImagePixelBox(sheet, image, box) {
  const left = Math.max(rowHeaderWidthForSheet(sheet), Math.round(box.left));
  const top = Math.max(COLUMN_HEADER_HEIGHT, Math.round(box.top));
  const width = Math.max(1, Math.round(box.width));
  const height = Math.max(1, Math.round(box.height));
  const anchor = nativeAnchorFromSheetPixel(sheet, "col", left);
  const rowAnchor = nativeAnchorFromSheetPixel(sheet, "row", top);
  const endCol = nativeAnchorFromSheetPixel(sheet, "col", left + width);
  const endRow = nativeAnchorFromSheetPixel(sheet, "row", top + height);
  ensureSheetSize(sheet, Math.max(1, endRow.index + 1), Math.max(1, endCol.index + 1));
  image.tl = {
    col: anchor.index,
    row: rowAnchor.index,
    nativeCol: anchor.index,
    nativeRow: rowAnchor.index,
    nativeColOff: anchor.offsetEmu,
    nativeRowOff: rowAnchor.offsetEmu,
  };
  image.br = null;
  image.ext = { width, height };
}



function ensureLineShapeBoxPadding(sheet, image) {
  if (!sheet || !isLineShapeObject(image)) return false;
  const box = imageBoxToPixels(sheet, image);
  if (!box) return false;
  const endpoints = lineEndpointPixelsFromBox(box, normalizeShapeLinePoints(image.shape?.points));
  const result = lineBoxAndPointsFromEndpoints(endpoints.start, endpoints.end, { shape: image.shape });
  if (boxesEqual(box, result.box) && linePointsEqual(image.shape?.points, result.points)) return false;
  setImagePixelBox(sheet, image, result.box);
  image.shape = {
    ...(image.shape || {}),
    points: normalizeShapeLinePoints(result.points),
  };
  return true;
}



function applySelectedShapeStyle(style, label) {
  const objects = selectedImageObjects();
  const pictureStyle = pictureStylePatch(style);
  const targets = objects.filter((image) => image?.shape || (!image?.shape && Object.keys(pictureStyle).length));
  if (!targets.length) {
    setStatus("図形または画像を選択してください。");
    return;
  }
  mutateShapeObjects(targets, (image) => {
    if (!image.shape) {
      applyPictureStylePatch(image, pictureStyle);
      return;
    }
    const nextShape = { ...(image.shape || {}), ...style };
    const richTextOverrideKeys = ["textColor", "textFontSize", "textBold", "textItalic", "textUnderline", "textFontFamily", "textOutlineColor", "textOutlineWidth"];
    if (richTextOverrideKeys.some((key) => Object.prototype.hasOwnProperty.call(style, key))) {
      delete nextShape.textRuns;
    }
    if (Object.prototype.hasOwnProperty.call(style, "fillColor") && !Object.prototype.hasOwnProperty.call(style, "fillGradient")) {
      delete nextShape.fillGradient;
    }
    if (Object.prototype.hasOwnProperty.call(style, "fillPattern") && style.fillPattern) {
      delete nextShape.fillGradient;
    } else if (Object.prototype.hasOwnProperty.call(style, "fillPattern")) {
      delete nextShape.fillPattern;
    }
    if (Object.prototype.hasOwnProperty.call(style, "fillColor") && !Object.prototype.hasOwnProperty.call(style, "fillPattern")) {
      delete nextShape.fillPattern;
    }
    image.shape = nextShape;
    refreshShapeSource(image);
  }, label, { objectLabel: "オブジェクト" });
}



function pictureStylePatch(style = {}) {
  const allowed = ["outlineColor", "outlineWidth", "lineDash", "lineCompound", "shadow", "glow", "softEdges", "reflection"];
  return Object.fromEntries(allowed.filter((key) => Object.prototype.hasOwnProperty.call(style, key)).map((key) => [key, style[key]]));
}



function applyPictureStylePatch(image, patch = {}) {
  Object.entries(patch).forEach(([key, value]) => {
    if (key === "outlineColor") image.outlineColor = normalizeColorValue(value) || "none";
    else if (key === "outlineWidth") image.outlineWidth = normalizeShapeOutlineWidth(value);
    else if (key === "lineDash") image.lineDash = normalizeShapeLineDash(value);
    else if (key === "lineCompound") image.lineCompound = normalizeShapeLineCompound(value);
    else if (key === "shadow") image.shadow = value || "";
    else if (key === "glow") image.glow = normalizeShapeGlow(value);
    else if (key === "softEdges") image.softEdges = normalizeShapeSoftEdges(value);
    else if (key === "reflection") image.reflection = normalizeShapeReflection(value);
  });
}



function applyShapeArrowPreset(value) {
  const preset = SHAPE_ARROW_PRESETS.find((item) => item.value === value) || SHAPE_ARROW_PRESETS[0];
  const images = selectedImageObjects().filter(isLineShapeObject);
  if (!images.length) {
    setStatus("矢印を設定する線を選択してください。");
    return;
  }
  mutateShapeObjects(images, (image, sheet) => {
    image.shape = {
      ...(image.shape || {}),
      lineStartArrow: preset.start,
      lineEndArrow: preset.end,
      lineStartArrowType: preset.start ? image.shape?.lineStartArrowType || "triangle" : "",
      lineEndArrowType: preset.end ? image.shape?.lineEndArrowType || "triangle" : "",
    };
    ensureLineShapeBoxPadding(sheet, image);
    refreshShapeSource(image);
  }, preset.label);
}



function applyShapeArrowTypePreset(side, value) {
  const normalizedSide = side === "start" ? "start" : "end";
  const arrowType = normalizeShapeArrowType(value);
  const images = selectedImageObjects().filter(isLineShapeObject);
  if (!images.length) {
    setStatus("矢印を設定する線を選択してください。");
    return;
  }
  mutateShapeObjects(images, (image, sheet) => {
    if (normalizedSide === "start") {
      image.shape = {
        ...(image.shape || {}),
        lineStartArrow: Boolean(arrowType),
        lineStartArrowType: arrowType,
      };
    } else {
      image.shape = {
        ...(image.shape || {}),
        lineEndArrow: Boolean(arrowType),
        lineEndArrowType: arrowType,
      };
    }
    ensureLineShapeBoxPadding(sheet, image);
    refreshShapeSource(image);
  }, normalizedSide === "start" ? "始点矢印の種類" : "終点矢印の種類");
}



function applyShapeArrowSizePreset(side, value) {
  const normalizedSide = side === "start" ? "start" : "end";
  const arrowSize = normalizeShapeArrowSize(value);
  const images = selectedImageObjects().filter(isLineShapeObject);
  if (!images.length) {
    setStatus("矢印を設定する線を選択してください。");
    return;
  }
  mutateShapeObjects(images, (image, sheet) => {
    if (normalizedSide === "start") {
      image.shape = {
        ...(image.shape || {}),
        lineStartArrow: true,
        lineStartArrowType: image.shape?.lineStartArrowType || "triangle",
        lineStartArrowSize: arrowSize,
      };
    } else {
      image.shape = {
        ...(image.shape || {}),
        lineEndArrow: true,
        lineEndArrowType: image.shape?.lineEndArrowType || "triangle",
        lineEndArrowSize: arrowSize,
      };
    }
    ensureLineShapeBoxPadding(sheet, image);
    refreshShapeSource(image);
  }, normalizedSide === "start" ? "始点矢印のサイズ" : "終点矢印のサイズ");
}



function applyShapeRotateAction(action) {
  const options = { updateConnectedLines: true };
  const objects = selectedImageObjects();
  if (!objects.length) {
    setStatus("回転するオブジェクトを選択してください。");
    return;
  }
  switch (action) {
    case "right90":
      mutateShapeObjects(objects, (image) => {
        setImageRotation(image, imageRotationDegrees(image) + 90);
        if (image.shape) refreshShapeSource(image);
      }, "右へ 90 度回転", { ...options, objectLabel: "オブジェクト" });
      break;
    case "left90":
      mutateShapeObjects(objects, (image) => {
        setImageRotation(image, imageRotationDegrees(image) - 90);
        if (image.shape) refreshShapeSource(image);
      }, "左へ 90 度回転", { ...options, objectLabel: "オブジェクト" });
      break;
    case "flipH":
      mutateShapeObjects(objects, (image) => {
        setImageFlip(image, "h", !imageFlipH(image));
      }, "左右反転", { ...options, objectLabel: "オブジェクト" });
      break;
    case "flipV":
      mutateShapeObjects(objects, (image) => {
        setImageFlip(image, "v", !imageFlipV(image));
      }, "上下反転", { ...options, objectLabel: "オブジェクト" });
      break;
    default:
      break;
  }
}



function addCustomShapeVertex(imageId, vertexKey) {
  const image = selectedImageObjectById(imageId);
  const vertex = customPathVertexForKey(image?.shape?.customPath, vertexKey);
  if (!image?.shape || image.shape.type !== CUSTOM_SHAPE_TYPE || !vertex) {
    setStatus("頂点を追加できる自由図形を選択してください。");
    return;
  }
  mutateShapeObjects([image], (target) => {
    const path = normalizeShapeCustomPath(target.shape?.customPath);
    const commands = path.commands.map((command) => ({ ...command, points: command.points.map((point) => ({ ...point })) }));
    const currentPoint = commands[vertex.commandIndex]?.points?.[vertex.pointIndex];
    const nextPoint = nextCustomPathAnchorPoint(commands, vertex.commandIndex);
    if (!currentPoint || !nextPoint) return;
    const midpoint = {
      x: (currentPoint.x + nextPoint.x) / 2,
      y: (currentPoint.y + nextPoint.y) / 2,
    };
    commands.splice(vertex.commandIndex + 1, 0, { cmd: "L", points: [midpoint] });
    target.shape = { ...(target.shape || {}), customPath: normalizeShapeCustomPath({ commands }) };
    refreshShapeSource(target);
  }, "頂点の追加");
  state.shapePointEditIds = [image.id];
}



function deleteCustomShapeVertex(imageId, vertexKey) {
  const image = selectedImageObjectById(imageId);
  const vertex = customPathVertexForKey(image?.shape?.customPath, vertexKey);
  if (!image?.shape || image.shape.type !== CUSTOM_SHAPE_TYPE || !vertex) {
    setStatus("削除する頂点を選択してください。");
    return;
  }
  if (vertex.command.cmd === "M") {
    setStatus("開始点は削除できません。");
    return;
  }
  if (customPathEditablePointCount(vertex.path) <= 2) {
    setStatus("これ以上頂点を削除できません。");
    return;
  }
  mutateShapeObjects([image], (target) => {
    const path = normalizeShapeCustomPath(target.shape?.customPath);
    const commands = path.commands.map((command) => ({ ...command, points: command.points.map((point) => ({ ...point })) }));
    const command = commands[vertex.commandIndex];
    if (!command) return;
    command.points.splice(vertex.pointIndex, 1);
    if (!command.points.length) commands.splice(vertex.commandIndex, 1);
    target.shape = { ...(target.shape || {}), customPath: normalizeShapeCustomPath({ commands }) };
    refreshShapeSource(target);
  }, "頂点の削除");
  state.shapePointEditIds = [image.id];
}



function toggleCustomShapePathClosed(imageId) {
  const image = selectedImageObjectById(imageId);
  const path = normalizeShapeCustomPath(image?.shape?.customPath);
  if (!image?.shape || image.shape.type !== CUSTOM_SHAPE_TYPE || !path) {
    setStatus("自由図形を選択してください。");
    return;
  }
  const closed = shapeCustomPathIsClosed(path);
  mutateShapeObjects([image], (target) => {
    const current = normalizeShapeCustomPath(target.shape?.customPath);
    const commands = current.commands
      .filter((command) => command.cmd !== "Z")
      .map((command) => ({ ...command, points: command.points.map((point) => ({ ...point })) }));
    if (!closed) commands.push({ cmd: "Z", points: [] });
    target.shape = { ...(target.shape || {}), customPath: normalizeShapeCustomPath({ commands }) };
    refreshShapeSource(target);
  }, closed ? "開いたパス" : "パスを閉じる");
  state.shapePointEditIds = [image.id];
}



function nextCustomPathAnchorPoint(commands, commandIndex) {
  for (let index = commandIndex + 1; index < commands.length; index += 1) {
    const command = commands[index];
    if (command?.points?.length) return command.points[command.points.length - 1];
  }
  return commands.find((command) => command?.cmd === "M" && command.points?.[0])?.points?.[0] || null;
}



function alignSelectedShapes(action) {
  const units = selectedShapeAlignmentUnits();
  const images = units.flatMap((unit) => unit.images);
  if (!units.length) {
    setStatus("配置するオブジェクトを選択してください。");
    return;
  }
  if (units.length < 2) {
    setStatus("整列するオブジェクトを 2 つ以上選択してください。");
    return;
  }
  const sheet = activeSheet();
  const boxes = units.filter((unit) => unit.box);
  const union = unionPixelBoxes(boxes.map((item) => item.box));
  if (!union) return;
  const sorted =
    action === "distribute-h"
      ? [...boxes].sort((a, b) => a.box.left - b.box.left)
      : action === "distribute-v"
        ? [...boxes].sort((a, b) => a.box.top - b.box.top)
        : boxes;
  const totalWidth = sorted.reduce((sum, item) => sum + item.box.width, 0);
  const totalHeight = sorted.reduce((sum, item) => sum + item.box.height, 0);
  const gapX = sorted.length > 1 ? (union.width - totalWidth) / (sorted.length - 1) : 0;
  const gapY = sorted.length > 1 ? (union.height - totalHeight) / (sorted.length - 1) : 0;
  let cursorX = union.left;
  let cursorY = union.top;
  const movements = [];
  sorted.forEach((unit) => {
    const { box } = unit;
    const next = { ...box };
    switch (action) {
      case "left":
        next.left = union.left;
        break;
      case "center":
        next.left = union.left + (union.width - box.width) / 2;
        break;
      case "right":
        next.left = union.left + union.width - box.width;
        break;
      case "top":
        next.top = union.top;
        break;
      case "middle":
        next.top = union.top + (union.height - box.height) / 2;
        break;
      case "bottom":
        next.top = union.top + union.height - box.height;
        break;
      case "distribute-h":
        next.left = cursorX;
        cursorX += box.width + gapX;
        break;
      case "distribute-v":
        next.top = cursorY;
        cursorY += box.height + gapY;
        break;
      default:
        break;
    }
    movements.push({ unit, dx: next.left - box.left, dy: next.top - box.top });
  });
  if (alignSelectedShapesForSheetGroup(images, movements, action)) return;
  const sheetBefore = captureSheetForHistory(sheet);
  movements.forEach((movement) => moveShapeAlignmentUnit(sheet, movement.unit, movement.dx, movement.dy));
  updateConnectedLinesForTargets(sheet, images.filter((image) => image && !isLineShapeObject(image)).map((image) => image.id));
  pruneDetachedLineShapeConnectionsForIds(sheet, images.filter((image) => isLineShapeObject(image)).map((image) => image.id));
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${selectedObjectSummary(images)} 配置`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`${selectedObjectSummary(images)}を配置しました。`);
}



function alignSelectedShapesForSheetGroup(images, movements, action) {
  const items = groupedShapeObjectTargets(images);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const moveBySourceId = new Map();
  (movements || []).forEach((movement) => {
    (movement.unit?.images || []).forEach((image) => {
      moveBySourceId.set(image.id, { dx: movement.dx, dy: movement.dy });
    });
  });
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      const movedShapeIds = [];
      const movedLineIds = [];
      item.images.forEach((target, index) => {
        const source = images[index];
        const movement = source ? moveBySourceId.get(source.id) : null;
        if (!movement) return;
        moveShapeAlignmentUnit(item.sheet, { images: [target] }, movement.dx, movement.dy);
        if (isLineShapeObject(target)) movedLineIds.push(target.id);
        else movedShapeIds.push(target.id);
      });
      updateConnectedLinesForTargets(item.sheet, movedShapeIds);
      pruneDetachedLineShapeConnectionsForIds(item.sheet, movedLineIds);
    });
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート${selectedObjectSummary(images)} 配置`,
    });
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
    setStatus(`${items.length} 個のシートの${selectedObjectSummary(images)}を配置しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクトを配置できませんでした：${error.message || error}`);
  }
  return true;
}



function selectedShapeAlignmentUnits() {
  const sheet = activeSheet();
  const images = selectedImageObjects();
  const used = new Set();
  const units = [];
  selectedImageGroups(images).forEach((group) => {
    const box = unionImageBox(sheet, group.images);
    if (!box) return;
    group.images.forEach((image) => used.add(image.id));
    units.push({ id: group.groupId, images: group.images, box });
  });
  images.forEach((image) => {
    if (used.has(image.id)) return;
    const box = imageBoxToPixels(sheet, image);
    if (box) units.push({ id: image.id, images: [image], box });
  });
  return units;
}



function moveShapeAlignmentUnit(sheet, unit, dx, dy) {
  unit.images.forEach((image) => {
    const box = imageBoxToPixels(sheet, image);
    if (!box) return;
    setImagePixelBox(sheet, image, {
      ...box,
      left: box.left + dx,
      top: box.top + dy,
    });
    if (image.shape) refreshShapeSource(image);
  });
}



function applySelectedShapeSize(axis, value) {
  if (!Number.isFinite(value) || value <= 0) {
    syncShapeFormatControls();
    return;
  }
  const objects = selectedImageObjects();
  if (!objects.length) {
    setStatus("サイズを変更するオブジェクトを選択してください。");
    return;
  }
  const group = selectedImageGroups(objects)[0];
  const isWholeGroupSelection = group && objects.length === group.images.length;
  if (isWholeGroupSelection) {
    resizeSelectedShapeGroup(axis, value, group);
    return;
  }
  mutateShapeObjects(objects, (image) => {
    image.ext = image.ext || {};
    image.ext[axis === "height" ? "height" : "width"] = Math.round(value);
    if (image.shape) refreshShapeSource(image);
  }, axis === "height" ? "高さの変更" : "幅の変更", { updateConnectedLines: true, objectLabel: "オブジェクト" });
}



function resizeSelectedShapeGroup(axis, value, group) {
  const sheet = activeSheet();
  if (resizeSelectedShapeGroupForSheetGroup(axis, value, group)) return;
  const groupBox = unionImageBox(sheet, group.images);
  if (!groupBox) {
    syncShapeFormatControls();
    return;
  }
  const targetSize = Math.max(1, Math.round(value));
  const scaleX = axis === "width" ? targetSize / Math.max(1, groupBox.width) : 1;
  const scaleY = axis === "height" ? targetSize / Math.max(1, groupBox.height) : 1;
  const sheetBefore = captureSheetForHistory(sheet);
  const movedLineIds = [];
	  group.images.forEach((image) => {
	    const box = imageBoxToPixels(sheet, image);
	    if (!box) return;
    setImagePixelBox(sheet, image, {
      left: groupBox.left + (box.left - groupBox.left) * scaleX,
      top: groupBox.top + (box.top - groupBox.top) * scaleY,
	      width: Math.max(1, box.width * scaleX),
	      height: Math.max(1, box.height * scaleY),
	    });
	    if (image.shape) refreshShapeSource(image);
    if (isLineShapeObject(image)) movedLineIds.push(image.id);
	  });
  updateConnectedLinesForTargets(sheet, group.images.filter((image) => image && !isLineShapeObject(image)).map((image) => image.id));
  pruneDetachedLineShapeConnectionsForIds(sheet, movedLineIds);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${group.groupName || "グループ"} ${axis === "height" ? "高さの変更" : "幅の変更"}`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`${group.groupName || "グループ"}の${axis === "height" ? "高さ" : "幅"}を変更しました。`);
}



function resizeSelectedShapeGroupForSheetGroup(axis, value, group) {
  if (!group?.images?.length) return false;
  const items = groupedShapeObjectTargets(group.images);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const targetSize = Math.max(1, Math.round(value));
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      const groupBox = unionImageBox(item.sheet, item.images);
      if (!groupBox) return;
      const scaleX = axis === "width" ? targetSize / Math.max(1, groupBox.width) : 1;
      const scaleY = axis === "height" ? targetSize / Math.max(1, groupBox.height) : 1;
      const movedLineIds = [];
      item.images.forEach((image) => {
        const box = imageBoxToPixels(item.sheet, image);
        if (!box) return;
        setImagePixelBox(item.sheet, image, {
          left: groupBox.left + (box.left - groupBox.left) * scaleX,
          top: groupBox.top + (box.top - groupBox.top) * scaleY,
          width: Math.max(1, box.width * scaleX),
          height: Math.max(1, box.height * scaleY),
        });
        if (image.shape) refreshShapeSource(image);
        if (isLineShapeObject(image)) movedLineIds.push(image.id);
      });
      updateConnectedLinesForTargets(item.sheet, item.images.filter((image) => image && !isLineShapeObject(image)).map((image) => image.id));
      pruneDetachedLineShapeConnectionsForIds(item.sheet, movedLineIds);
    });
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート${group.groupName || "グループ"} ${axis === "height" ? "高さの変更" : "幅の変更"}`,
    });
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
    setStatus(`${items.length} 個のシートの${group.groupName || "グループ"}の${axis === "height" ? "高さ" : "幅"}を変更しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`グループ サイズを変更できませんでした：${error.message || error}`);
  }
  return true;
}



function resetSelectedShapeSize() {
  mutateSelectedShape((image) => {
    const descriptor = shapeDescriptorForType(image.shape?.type || DEFAULT_SHAPE_TYPE, image.shape || {});
    image.ext = { width: descriptor.width, height: descriptor.height };
  }, "サイズのリセット", { updateConnectedLines: true });
}



function applySelectedShapeRotation(value) {
  const objects = selectedImageObjects();
  if (!objects.length) {
    setStatus("回転するオブジェクトを選択してください。");
    return;
  }
  mutateShapeObjects(objects, (image) => {
    setImageRotation(image, value);
    if (image.shape) refreshShapeSource(image);
  }, "回転", { updateConnectedLines: true, objectLabel: "オブジェクト" });
}



function refreshShapeSource(image) {
  if (image?.chart?.series?.length && (!image.shape || image.shape.sourcePreset === "graphicFrame")) {
    const width = Math.max(1, Math.round(Number(image.ext?.width) || 520));
    const height = Math.max(1, Math.round(Number(image.ext?.height) || 320));
    image.src = svgDataUrl(chartSvg(image.chart, image.chart.type || "column", { width, height, title: image.chart.title || image.name || "グラフ" }));
    image.mimeType = "image/svg+xml";
    image.extension = "svg";
    image.name = image.name || image.chart.title || "グラフ";
    return;
  }
  if (!image?.shape) return;
  const descriptor = shapeDescriptorForType(image.shape.type || DEFAULT_SHAPE_TYPE, {
    ...image.shape,
    width: image.ext?.width,
    height: image.ext?.height,
    includeTextInSvg: false,
  });
  image.src = descriptor.src;
  image.mimeType = descriptor.mimeType;
  image.extension = descriptor.extension;
  image.name = image.name || image.shape.label || descriptor.name || "図形";
  image.shape = { ...image.shape, ...descriptor.shape };
}



function reorderSelectedShape(direction) {
  if (!selectedImageObjects().length) {
    setStatus("オブジェクトを選択してください。");
    return;
  }
  reorderSelectedObjects(direction);
}



function groupSelectedShapes() {
  const objects = selectedImageObjects();
  if (objects.length < 2) {
    setStatus("グループ化するオブジェクトを 2 つ以上選択してください。");
    return;
  }
  if (groupSelectedShapesForSheetGroup(objects)) return;
  const sheet = activeSheet();
  const sheetBefore = captureSheetForHistory(sheet);
  const groupId = uniqueWorksheetObjectId("group");
  const groupName = `グループ ${nextShapeGroupNumber(sheet)}`;
  objects.forEach((image) => {
    image.groupId = groupId;
    image.groupName = groupName;
  });
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${groupName} グループ化`,
  });
  setSelectedImageIds(objects.map((image) => image.id), objects[0].id);
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`${objects.length} 個のオブジェクトをグループ化しました。`);
}



function groupSelectedShapesForSheetGroup(objects) {
  const items = groupedShapeObjectTargets(objects);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      const groupId = uniqueWorksheetObjectId("group");
      const groupName = `グループ ${nextShapeGroupNumber(item.sheet)}`;
      item.images.forEach((image) => {
        image.groupId = groupId;
        image.groupName = groupName;
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
      label: `${items.length} シート グループ化`,
    });
    setSelectedImageIds(objects.map((image) => image.id), objects[0].id);
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
    setStatus(`${items.length} 個のシートで${objects.length} 個のオブジェクトをグループ化しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクトをグループ化できませんでした：${error.message || error}`);
  }
  return true;
}



function ungroupSelectedShapes() {
  const objects = selectedImageObjects().filter((image) => image?.groupId);
  if (!objects.length) {
    setStatus("グループ解除するオブジェクトを選択してください。");
    return;
  }
  if (ungroupSelectedShapesForSheetGroup(objects)) return;
  const sheet = activeSheet();
  const sheetBefore = captureSheetForHistory(sheet);
  const ids = objects.map((image) => image.id);
  objects.forEach((image) => {
    delete image.groupId;
    delete image.groupName;
  });
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${selectedObjectSummary(objects)} グループ解除`,
  });
  setSelectedImageIds(ids, ids[0]);
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`${objects.length} 個のオブジェクトをグループ解除しました。`);
}



function ungroupSelectedShapesForSheetGroup(objects) {
  const items = groupedShapeObjectTargets(objects);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      item.images.forEach((image) => {
        delete image.groupId;
        delete image.groupName;
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
      label: `${items.length} シート${selectedObjectSummary(objects)} グループ解除`,
    });
    setSelectedImageIds(objects.map((image) => image.id), objects[0].id);
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
    setStatus(`${items.length} 個のシートで${objects.length} 個のオブジェクトをグループ解除しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`グループ解除できませんでした：${error.message || error}`);
  }
  return true;
}



function nextShapeGroupNumber(sheet) {
  const max = (sheet.images || []).reduce((value, image) => {
    const match = String(image.groupName || "").match(/グループ\s+(\d+)/);
    return match ? Math.max(value, Number(match[1]) || 0) : value;
  }, 0);
  return max + 1;
}



function reorderSelectedObjects(direction) {
  const objects = selectedImageObjects();
  if (!objects.length) {
    setStatus("オブジェクトを選択してください。");
    return;
  }
  const sheet = activeSheet();
  if (reorderSelectedObjectsForSheetGroup(objects, direction)) return;
  const selectedIds = new Set(objects.map((image) => image.id));
  const images = sheet.images || [];
  const nextImages = reorderedImagesForDirection(sheet, images, selectedIds, direction);
  if (!nextImages) return;
  if (nextImages.every((image, index) => image === images[index])) return;
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.images = nextImages;
  const directionLabel = {
    front: "最前面へ移動",
    forward: "前面へ移動",
    backward: "背面へ移動",
    back: "最背面へ移動",
  }[direction] || "順序変更";
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${selectedObjectSummary(objects)} ${directionLabel}`,
  });
  renderSheet();
  updateSelectionUi();
  setStatus(`${selectedObjectSummary(objects)}を${directionLabel}しました。`);
}



function reorderedImagesForDirection(sheet, images, selectedIds, direction) {
  const selected = (images || []).filter((image) => selectedIds.has(image.id));
  const others = (images || []).filter((image) => !selectedIds.has(image.id));
  if (!selected.length) return null;
  if (direction === "front") return [...others, ...selected];
  if (direction === "back") return [...selected, ...others];

  const overlapCandidate = nearestOverlappingLayerCandidate(sheet, images, selectedIds, direction);
  if (overlapCandidate) {
    return moveSelectedLayerBlockAroundCandidate(images, selectedIds, overlapCandidate.id, direction);
  }

  return reorderImagesByAdjacentLayer(images, selectedIds, direction);
}



function reorderImagesByAdjacentLayer(images, selectedIds, direction) {
  const nextImages = [...(images || [])];
  const layerIndexes = nextImages
    .map((image, index) => (!image.hidden || selectedIds.has(image.id) ? index : -1))
    .filter((index) => index >= 0);
  if (direction === "forward") {
    for (let layer = layerIndexes.length - 2; layer >= 0; layer -= 1) {
      const index = layerIndexes[layer];
      const nextIndex = layerIndexes[layer + 1];
      if (selectedIds.has(nextImages[index].id) && !selectedIds.has(nextImages[nextIndex].id)) {
        [nextImages[index], nextImages[nextIndex]] = [nextImages[nextIndex], nextImages[index]];
      }
    }
    return nextImages;
  }
  if (direction === "backward") {
    for (let layer = 1; layer < layerIndexes.length; layer += 1) {
      const index = layerIndexes[layer];
      const previousIndex = layerIndexes[layer - 1];
      if (selectedIds.has(nextImages[index].id) && !selectedIds.has(nextImages[previousIndex].id)) {
        [nextImages[previousIndex], nextImages[index]] = [nextImages[index], nextImages[previousIndex]];
      }
    }
    return nextImages;
  }
  return null;
}



function nearestOverlappingLayerCandidate(sheet, images, selectedIds, direction) {
  if (direction !== "forward" && direction !== "backward") return null;
  const list = images || [];
  const selectedEntries = list
    .map((image, index) => (selectedIds.has(image.id) ? { image, index, bounds: imageBoundsForLayerReorder(sheet, image) } : null))
    .filter((entry) => entry?.bounds);
  if (!selectedEntries.length) return null;
  const selectedBounds = selectedEntries.map((entry) => entry.bounds);
  const selectedIndexes = selectedEntries.map((entry) => entry.index);
  if (direction === "forward") {
    const topIndex = Math.max(...selectedIndexes);
    for (let index = topIndex + 1; index < list.length; index += 1) {
      const candidate = list[index];
      if (!candidate || candidate.hidden || selectedIds.has(candidate.id)) continue;
      if (imageBoundsOverlapAnySelected(sheet, candidate, selectedBounds)) return candidate;
    }
  } else {
    const bottomIndex = Math.min(...selectedIndexes);
    for (let index = bottomIndex - 1; index >= 0; index -= 1) {
      const candidate = list[index];
      if (!candidate || candidate.hidden || selectedIds.has(candidate.id)) continue;
      if (imageBoundsOverlapAnySelected(sheet, candidate, selectedBounds)) return candidate;
    }
  }
  return null;
}



function imageBoundsForLayerReorder(sheet, image) {
  const box = imageBoxToPixels(sheet, image);
  return box ? imageVisualBoundsToPixels(image, box) : null;
}



function imageBoundsOverlapAnySelected(sheet, image, selectedBounds) {
  const bounds = imageBoundsForLayerReorder(sheet, image);
  return Boolean(bounds && selectedBounds.some((selectedBox) => pixelBoxesIntersect(bounds, selectedBox)));
}



function moveSelectedLayerBlockAroundCandidate(images, selectedIds, candidateId, direction) {
  const selected = (images || []).filter((image) => selectedIds.has(image.id));
  const others = (images || []).filter((image) => !selectedIds.has(image.id));
  const candidateIndex = others.findIndex((image) => image.id === candidateId);
  if (candidateIndex < 0) return null;
  const insertIndex = direction === "forward" ? candidateIndex + 1 : candidateIndex;
  return [...others.slice(0, insertIndex), ...selected, ...others.slice(insertIndex)];
}



function reorderSelectedObjectsForSheetGroup(objects, direction) {
  const items = groupedShapeObjectTargets(objects);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const directionLabel = {
    front: "最前面へ移動",
    forward: "前面へ移動",
    backward: "背面へ移動",
    back: "最背面へ移動",
  }[direction] || "順序変更";
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  let changed = false;
  try {
    items.forEach((item) => {
      const images = item.sheet.images || [];
      const selectedIds = new Set(item.images.map((image) => image.id));
      const nextImages = reorderedImagesForDirection(item.sheet, images, selectedIds, direction);
      if (!nextImages || nextImages.every((image, index) => image === images[index])) return;
      item.sheet.images = nextImages;
      changed = true;
    });
    if (!changed) return true;
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート${selectedObjectSummary(objects)} ${directionLabel}`,
    });
    renderSheet();
    updateSelectionUi();
    setStatus(`${items.length} 個のシートの${selectedObjectSummary(objects)}を${directionLabel}しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクトの順序を変更できませんでした：${error.message || error}`);
  }
  return true;
}



function mutateSelectedShape(mutator, label, options = {}) {
  const images = selectedImageObjects().filter((image) => image?.shape);
  if (!images.length) {
    setStatus("図形を選択してください。");
    return;
  }
  mutateShapeObjects(images, mutator, label, options);
}



function mutateShapeObjectsForSheetGroup(images, mutator, label, options = {}) {
  const items = groupedShapeObjectTargets(images);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      item.images.forEach((image) => mutator(image, item.sheet));
      if (options.updateConnectedLines) {
        updateConnectedLinesForTargets(item.sheet, item.images.filter((image) => image && !isLineShapeObject(image)).map((image) => image.id));
        pruneDetachedLineShapeConnectionsForIds(item.sheet, item.images.filter((image) => isLineShapeObject(image)).map((image) => image.id));
      }
    });
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート${selectedObjectSummary(images)} ${label}`,
    });
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
    setStatus(`${items.length} 個のシートの${selectedObjectSummary(images)}に${label}を適用しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクトを編集できませんでした：${error.message || error}`);
  }
  return true;
}

export {
  addCustomShapeVertex,
  alignSelectedShapes,
  alignSelectedShapesForSheetGroup,
  applyPictureStylePatch,
  applySelectedShapeRotation,
  applySelectedShapeSize,
  applySelectedShapeStyle,
  applyShapeArrowPreset,
  applyShapeArrowSizePreset,
  applyShapeArrowTypePreset,
  applyShapeRotateAction,
  blockArrowConnectionSites,
  boxesEqual,
  calloutConnectionSites,
  cardinalConnectionSites,
  compassConnectionSites,
  connectionCandidateMatchesSite,
  connectionSite,
  constrainImageTransformAspect,
  currentPictureCropForImage,
  customShapeConnectionSites,
  deleteCustomShapeVertex,
  ensureLineShapeBoxPadding,
  enterShapePointEditMode,
  finishPictureCropDrag,
  finishSelectedImageTransform,
  finishSelectedImageTransformForSheetGroup,
  fitViewBoxPoints,
  groupSelectedShapes,
  groupSelectedShapesForSheetGroup,
  imageBoundsForLayerReorder,
  imageBoundsOverlapAnySelected,
  imageTransformBox,
  imageTransformCustomPathPoint,
  imageTransformGroupResizeBoxes,
  imageTransformGroupRotationBoxes,
  imageTransformGroupVisualBox,
  imageTransformLinePoint,
  imageTransformMoveBoxes,
  imageTransformMoveDelta,
  imageTransformRotation,
  imageTransformShapeAdjustment,
  insertCopiedImagesFromMoveTransform,
  isCornerResizeHandle,
  lineConnectionForGroupedTarget,
  lineConnectionPreviewTarget,
  lineConnectionTargetObject,
  lineEndpointPixelsFromBox,
  linePointsEqual,
  lineShapePointerHitsStroke,
  lineShapeRoutePointsLocal,
  lineShapeSegmentsInSheetPixels,
  maybeSnapImageBoxTopLeftToGrid,
  moveSelectedImagesBy,
  moveSelectedImagesBySheetGroup,
  moveSelectedLayerBlockAroundCandidate,
  moveShapeAlignmentUnit,
  mutateSelectedShape,
  mutateShapeObjectsForSheetGroup,
  nearestOverlappingLayerCandidate,
  nextCustomPathAnchorPoint,
  nextShapeGroupNumber,
  normalizeInteractivePictureCrop,
  pictureCropFullBoxFromVisibleBox,
  pictureCropVisibleBoxFromFullBox,
  pictureStylePatch,
  pointToSegmentDistance,
  polygonConnectionSites,
  polygonPointAliases,
  previewConnectedLinesForTransform,
  previewSelectedImageBox,
  previewSelectedImageGroupFrame,
  primaryImagePositionElement,
  pruneDetachedLineShapeConnections,
  pruneDetachedLineShapeConnectionsForIds,
  refreshShapeSource,
  regularPolygonConnectionSites,
  remapPastedLineShapeConnections,
  removeLineConnectionPreviewOverlay,
  reorderImagesByAdjacentLayer,
  reorderSelectedObjects,
  reorderSelectedObjectsForSheetGroup,
  reorderSelectedShape,
  reorderedImagesForDirection,
  resetSelectedShapeSize,
  resizeSelectedShapeGroup,
  resizeSelectedShapeGroupForSheetGroup,
  rotatePointAround,
  sameRotation,
  sampleCubicBezierPoints,
  selectedShapeAlignmentUnits,
  selectionIdsForImage,
  sequentialConnectionSites,
  setImageFlip,
  setImagePixelBox,
  setImageRotation,
  setLineShapeConnections,
  shapeAdjustmentFromPointer,
  shapeAdjustmentsEqual,
  shapeConnectionPointCandidates,
  shapeConnectionPointForSite,
  shapeConnectionSitesForImage,
  shapeCustomPathsEqual,
  shouldPreservePictureAspectOnResize,
  snapLineEndpointToShape,
  snapPixelToNearestGridLine,
  snapRotationDelta,
  startPictureCropDrag,
  startSelectedImageTransform,
  svgPointsStringToNormalizedPoints,
  toggleCustomShapePathClosed,
  togglePictureCropMode,
  transformConnectionBoxOverrides,
  transformConnectionRotationOverrides,
  transformConnectionTargetIds,
  transformedImageBoxForSheetTarget,
  ungroupSelectedShapes,
  ungroupSelectedShapesForSheetGroup,
  unionPixelBoxes,
  updateConnectedLinesForTargets,
  updateLineConnectionPreviewOverlay,
  updateLineHandlePositions,
  updateLineShapeConnection,
  updatePictureCropDrag,
  updateSelectedImageTransform,
  updateShapeAdjustmentHandlePositions,
};
