import { reconcileViewportGrid } from '../../../src/modules/rendering/viewport-reconciliation.js';

export async function checkViewportContinuity(document) {
  const results = [];
  const assert = (ok, label) => { if (!ok) throw new Error(label); };
  const host = document.createElement('div');
  document.body.append(host);
  const svg = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="80" height="30"><rect width="80" height="30" fill="teal"/></svg>');
  const drawing = (id, text = 'Layout', extra = '') => `<div class="sheet-image sheet-shape" data-image-id="${id}" ${extra}><img src="${svg}" alt=""><span>${text}</span></div>`;
  const picture = `<img class="sheet-image" data-image-id="picture" src="${svg}" alt="Picture">`;
  const linked = `<a class="sheet-image-link" data-image-id="linked" href="#fixture"><div class="sheet-image sheet-picture-crop" data-image-id="linked"><img src="${svg}" alt="Cropped"></div></a>`;
  const grid = (content, name = 'Layout') => {
    const template = document.createElement('template');
    template.innerHTML = `<div class="sheet-grid" data-sheet="${name}" style="height:300px">${content}</div>`;
    return template.content.firstElementChild;
  };
  const run = async (name, test) => { await test(); results.push({ name, passed: true }); };
  try {
    await run('Old full replacement reproduces detached decoded drawings', () => {
      host.replaceChildren(grid(drawing('shape')));
      const previous = host.querySelector('[data-image-id]');
      host.innerHTML = grid(drawing('shape')).outerHTML;
      assert(!previous.isConnected, 'baseline must reproduce detachment');
    });
    await run('120 viewport frames retain connected shape, picture and linked/cropped images', async () => {
      const current = grid(drawing('shape') + picture + linked);
      host.replaceChildren(current);
      const drawings = Array.from(current.children);
      const images = Array.from(current.querySelectorAll('img'));
      await Promise.all(images.map(img => img.decode()));
      const observer = new MutationObserver(() => {});
      observer.observe(current, { childList: true, subtree: true, attributes: true });
      for (let frame = 0; frame < 120; frame++) {
        const next = grid(`<div class="cell">Row ${frame}</div>` + drawing('shape') + `<div class="cell">Column ${frame}</div>` + picture + linked + (frame % 2 ? drawing('edge') : ''));
        const expectedMarkup = next.outerHTML;
        assert(reconcileViewportGrid(current, next), 'reconciliation rejected');
        assert(drawings.every(node => node.isConnected && node.parentNode === current), 'drawing disconnected');
        assert(images.every(img => img.isConnected && img.complete && img.naturalWidth > 0), 'decoded image lost');
        for (const record of observer.takeRecords()) {
          assert(!Array.from(record.removedNodes).some(node => drawings.includes(node)), 'retained node removed');
          assert(!(record.type === 'attributes' && images.includes(record.target)), 'retained image attributes reset');
        }
        assert(current.outerHTML === expectedMarkup, 'viewport content/order differs');
        await new Promise(requestAnimationFrame);
      }
      observer.disconnect();
    });
    await run('Content/style edits replace changed shape without stale text', () => {
      const current = host.firstElementChild;
      const old = current.querySelector('[data-image-id="shape"]');
      reconcileViewportGrid(current, grid(drawing('shape', 'Updated', 'style="color:red"') + picture));
      assert(!old.isConnected && current.textContent === 'Updated', 'edit not applied');
    });
    await run('Leaving/reentering viewport removes stale drawings and restores content', () => {
      const current = host.firstElementChild;
      reconcileViewportGrid(current, grid('<div>Only cells</div>'));
      assert(!current.querySelector('[data-image-id]'), 'stale drawing retained');
      reconcileViewportGrid(current, grid(drawing('shape')));
      assert(current.querySelector('[data-image-id="shape"]').isConnected, 'drawing not restored');
    });
    await run('Reordered drawings preserve requested stacking order', () => {
      const current = host.firstElementChild;
      reconcileViewportGrid(current, grid(drawing('a') + drawing('b')));
      reconcileViewportGrid(current, grid(drawing('b') + drawing('a')));
      assert(current.children[0].dataset.imageId === 'b', 'wrong z order');
    });
    await run('Duplicate IDs cannot reuse the wrong drawing', () => {
      const current = host.firstElementChild;
      const old = current.firstElementChild;
      reconcileViewportGrid(current, grid(drawing('b', 'First') + drawing('b', 'Second')));
      assert(!old.isConnected && current.children.length === 2, 'ambiguous drawing reused');
      assert(current.textContent === 'FirstSecond', 'duplicate content lost');
    });
    await run('Grid attributes update and sheet switches request a full render', () => {
      const current = host.firstElementChild;
      current.setAttribute('data-obsolete', 'true');
      const next = grid(picture); next.style.height = '450px';
      reconcileViewportGrid(current, next);
      assert(current.style.height === '450px' && !current.hasAttribute('data-obsolete'), 'stale grid attributes');
      assert(!reconcileViewportGrid(current, grid(picture, 'Other sheet')), 'sheet boundary ignored');
      assert(!reconcileViewportGrid(null, next), 'missing grid boundary ignored');
    });
  } finally { host.remove(); }
  return results;
}
