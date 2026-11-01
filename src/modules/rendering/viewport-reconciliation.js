// Keep decoded drawings attached while the virtualized cells around them change.
// Moving a preserved image through a detached fragment still causes Safari to
// repaint/redecode it, so normal scrolling must leave that node in place.
function viewportDrawingKey(node) {
  if (!node.matches?.('.sheet-image, .sheet-image-link')) return null;
  return node.getAttribute('data-image-id') || null;
}

function uniqueViewportDrawings(grid) {
  const drawings = new Map();
  for (const node of grid.children) {
    const key = viewportDrawingKey(node);
    if (key) drawings.set(key, drawings.has(key) ? null : node);
  }
  return drawings;
}

function reconcileViewportGrid(current, next) {
  if (!current || !next || current.tagName !== next.tagName
    || current.getAttribute('data-sheet') !== next.getAttribute('data-sheet')) return false;

  const existing = uniqueViewportDrawings(current);
  const incoming = uniqueViewportDrawings(next);
  const preserved = new Set();
  const children = Array.from(next.childNodes, (node) => {
    const key = viewportDrawingKey(node);
    const previous = key && incoming.get(key) === node ? existing.get(key) : null;
    if (!previous || !previous.isEqualNode(node)) return node;
    preserved.add(previous);
    return previous;
  });

  for (const attribute of Array.from(current.attributes)) {
    if (!next.hasAttribute(attribute.name)) current.removeAttribute(attribute.name);
  }
  for (const attribute of next.attributes) {
    if (current.getAttribute(attribute.name) !== attribute.value) current.setAttribute(attribute.name, attribute.value);
  }
  for (const node of Array.from(current.childNodes)) {
    if (!preserved.has(node)) node.remove();
  }
  let cursor = current.firstChild;
  for (const node of children) {
    if (node === cursor) cursor = cursor.nextSibling;
    else current.insertBefore(node, cursor);
  }
  return true;
}

export { reconcileViewportGrid };
