import assert from "node:assert/strict";
import { test } from "node:test";
import { gzipSync } from "node:zlib";
import { chromium } from "@playwright/test";
import { preview } from "vite";

test("saved HTML and refreshed browser storage render compressed and legacy pictures", async (t) => {
  const server = await preview({
    logLevel: "error",
    preview: { host: "127.0.0.1", port: 0 },
    build: { outDir: process.env.CWS_DIST_DIR || "dist" },
  });
  t.after(() => server.close());
  const origin = `http://127.0.0.1:${server.httpServer.address().port}`;
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  await page.route(/^https?:/, (route) =>
    new URL(route.request().url()).origin === origin ? route.continue() : route.abort());
  await page.goto(origin);

  const png = "iVBORw0KGgoAAAANSUhEUgAAAAUAAAADCAIAAADUVFKvAAAAE0lEQVR4nGNUztzCgASYkDlY+AA5rwFGo/pAOgAAAABJRU5ErkJggg==";
  const zipped = gzipSync(Buffer.from(png, "base64")).toString("base64");
  const asset = { data: zipped, mimeType: "image/png" };
  const sources = [
    { assetId: "marked" }, { assetId: "legacy" }, { assetId: "legacy" },
    { src: `data:image/png;base64,${zipped}` }, { src: `data:image/png;base64,${png}` },
  ];
  const images = sources.map((source, i) => ({
    ...source, id: `pic-${i}`, name: `picture-${i}.png`,
    tl: { row: i, col: 0, nativeRow: i, nativeCol: 0 }, ext: { width: 40, height: 24 },
  }));
  const model = {
    assets: { images: { marked: { ...asset, dataEncoding: "gzip-base64" }, legacy: asset } },
    sheets: [{ name: "Pictures", images }],
  };
  const html = `<script id="websheet-model" type="application/json">${JSON.stringify(model)}</script>`;
  await page.setInputFiles("#fileInput", { name: "pictures.html", mimeType: "text/html", buffer: Buffer.from(html) });

  async function assertRendered() {
    for (const picture of images) {
      const img = page.locator(`img[data-image-id="${picture.id}"]`);
      await img.waitFor({ state: "visible" });
      const pixels = await img.evaluate(async (node) => {
        await node.decode();
        const canvas = document.createElement("canvas");
        const ctx = canvas.getContext("2d");
        ctx.drawImage(node, 0, 0);
        return [node.naturalWidth, node.naturalHeight, ...ctx.getImageData(0, 0, 1, 1).data];
      });
      assert.deepEqual(pixels, [5, 3, 35, 105, 180, 255], picture.name);
    }
  }
  await assertRendered();
  await page.waitForFunction(() => document.querySelector("#fileInput")?.files.length === 0);
  await page.reload();
  await assertRendered();
});
