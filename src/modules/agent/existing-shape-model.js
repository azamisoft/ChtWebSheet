// Shared by the host and agent bundle. Revisions bind edits to the exact
// observed object; they are stale-write guards, not authentication tokens.
function cwsShapeStyleKeys() { return new Set(['textAlign','textVAlign','textMargin','textFontSize','textFontFamily','textColor','textBold','textItalic','textUnderline','fillColor','outlineColor','outlineWidth']); }

function canonical(value) {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === 'object') return Object.fromEntries(Object.keys(value).sort().filter(key => value[key] !== undefined).map(key => [key,canonical(value[key])]));
  return value;
}

function cwsShapeRevision(image) {
  const snapshot = Object.fromEntries(['id','name','tl','br','ext','shape','cwsLayout'].map(key => [key,image[key]]));
  let hash = 0xcbf29ce484222325n;
  for (const byte of new TextEncoder().encode(JSON.stringify(canonical(snapshot)))) hash = BigInt.asUintN(64,(hash ^ BigInt(byte)) * 0x100000001b3n);
  return 'shape-v1-' + hash.toString(16).padStart(16,'0');
}

function cwsNormalizeShapeStylePatch(style) {
  if (!style || typeof style !== 'object' || Array.isArray(style) || !Object.keys(style).length) throw new Error('A nonempty sparse shape style is required.');
  const result = {};
  for (const [key,value] of Object.entries(style)) {
    if (!cwsShapeStyleKeys().has(key)) throw new Error('Unsupported shape style field: ' + key);
    if (key === 'textAlign' && !['left','center','right','justify','distributed'].includes(value)) throw new Error('Invalid text alignment.');
    if (key === 'textVAlign' && !['top','middle','bottom'].includes(value)) throw new Error('Invalid vertical alignment.');
    if (key === 'textMargin') {
      if (!value || typeof value !== 'object' || Array.isArray(value) || !Object.keys(value).length || Object.entries(value).some(([side,n]) => !['left','right','top','bottom'].includes(side) || !Number.isFinite(n) || n<0 || n>256)) throw new Error('Invalid text margins.');
      result[key]={...value};continue;
    }
    if (['textFontSize','outlineWidth'].includes(key) && (!Number.isFinite(value) || value < (key === 'textFontSize' ? 6 : 0) || value > 96)) throw new Error('Invalid shape style size.');
    if (['textBold','textItalic'].includes(key) && typeof value !== 'boolean') throw new Error('Invalid text style flag.');
    if (key === 'textUnderline' && ![true,false,'','none','underline'].includes(value)) throw new Error('Invalid underline.');
    if (['textColor','fillColor','outlineColor'].includes(key) && (typeof value !== 'string' || !/^(#[\da-f]{6}|none|transparent)$/i.test(value))) throw new Error('Invalid shape color.');
    if (key === 'textFontFamily' && (typeof value !== 'string' || !value.trim())) throw new Error('Invalid font family.');
    result[key]=value;
  }
  return result;
}

function cwsNormalizeExistingShapeEdit(op) {
  const type=String(op.type || '').toLowerCase()==='removeshape'?'removeShape':'setShapeStyle';
  if (!['removeshape','setshapestyle'].includes(String(op.type || '').toLowerCase()) || typeof op.shapeId !== 'string' || !op.shapeId || !/^shape-v1-[\da-f]{16}$/.test(op.expectedRevision || '')) throw new Error('Exact shape ID and observed revision are required.');
  const result={type,shapeId:op.shapeId,expectedRevision:op.expectedRevision};
  if (type==='setShapeStyle') result.style=cwsNormalizeShapeStylePatch(op.style);
  return result;
}

function cwsExistingShapeEditTarget(sheet,op) {
  const matches=(sheet?.images || []).filter(image => image.id===op.shapeId && image.shape);
  if (matches.length!==1 || cwsShapeRevision(matches[0])!==op.expectedRevision) throw new Error('The shape changed since it was read. Read the current object again.');
  const image=matches[0];
  if (op.type==='removeShape') {
    if (image.cwsLayout?.componentId || image.cwsLayout?.componentLabelShape) throw new Error('Removing an identified component requires a coordinated catalogue edit.');
    const referenced=(sheet.images || []).some(other => other.id!==image.id && ['parentShape','componentLabelShape','fromShape','toShape'].some(key => other.cwsLayout?.[key]===image.name || other.shape?.[key]===image.name));
    if (referenced) throw new Error('A referenced shape or a container with children cannot be removed in isolation.');
  }
  return image;
}

function cwsApplyExistingShapeEdit(sheet,op) {
  const image=cwsExistingShapeEditTarget(sheet,op);
  if (op.type==='removeShape') sheet.images=sheet.images.filter(other => other.id!==image.id);
  else {
    const style=cwsNormalizeShapeStylePatch(op.style);
    if (style.textMargin) style.textMargin={...(image.shape.textMargin || {}),...style.textMargin};
    image.shape={...image.shape,...style};
  }
  return image;
}

function cwsExistingShapeEvidence(image,bounds) {
  const style=Object.fromEntries(Object.entries(image.shape || {}).filter(([key]) => cwsShapeStyleKeys().has(key) || ['textRuns','points','lineStartArrow','lineEndArrow','lineDash','rotation','flipH','flipV','fromShape','toShape'].includes(key)));
  return {shapeId:image.id,name:image.name,text:image.shape?.text || '',type:image.shape?.type,revision:cwsShapeRevision(image),style,layout:Object.fromEntries(Object.entries(image.cwsLayout || {}).filter(([,value])=>value!==null&&value!==undefined)),bounds};
}

export { cwsShapeRevision, cwsNormalizeShapeStylePatch, cwsNormalizeExistingShapeEdit, cwsExistingShapeEditTarget, cwsApplyExistingShapeEdit, cwsExistingShapeEvidence };
