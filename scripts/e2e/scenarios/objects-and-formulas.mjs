// Sequential E2E feature module.
import {
  STANDALONE_RUNTIME_CHANNEL,
  WEBSHEET_MODEL_SCRIPT_RE,
  path,
  pathToFileURL,
  readFile,
  workbookPath,
  writeFile,
} from "../../e2e-real-upload.mjs";
import {
  exportedStandaloneModelFromHtml,
  looksLikeStandaloneWorkbookHtml,
  replaceStandaloneModelPayload,
  replaceStandaloneModelScriptRaw,
} from "../helpers/fixtures-and-model.mjs";
async function runObjectsAndFormulas(scenarioState) {
  let { context, page, messages, result, exportedHtmlFilePath, exportedHtml, runtimeBaseUrl, runtimeSource, saveAsDialog } = scenarioState;

  await page.click('[data-ribbon-tab="file"]');


  await page.waitForFunction(() => !document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 10000 });


  await page.click('#fileSideMenu [data-ribbon-command="save-as"]');


  await page.waitForSelector('[data-data-dialog="save-as"]', { timeout: 10000 });


  await page.selectOption("[data-data-field='saveAsType']", "xlsx");


  await page.click('[data-data-dialog-action="save-as-browse"]');


  await page.waitForTimeout(500);


  saveAsDialog.handlePicked = await page.evaluate(() => ({
    location: document.querySelector("[data-data-field='saveAsLocation']")?.value || "",
    fileName: document.querySelector("[data-data-field='saveAsFileName']")?.value || "",
    note: document.querySelector("[data-save-as-note]")?.textContent || "",
    pickerCalls: window.__webSheetSavePickerCalls || [],
    dialogOpen: Boolean(document.querySelector("[data-data-dialog='save-as']")),
    browseButtonExists: Boolean(document.querySelector("[data-data-dialog-action='save-as-browse']")),
  }));


  if (!saveAsDialog.handlePicked.location?.includes("選択済み")) {
    throw new Error(`Save picker did not update the Save As dialog: ${JSON.stringify(saveAsDialog.handlePicked)}`);
  }


  await page.click('[data-data-dialog-action="save-as-ok"]');


  await page.waitForFunction(() => (window.__webSheetSaveHandleWrites || []).filter((entry) => entry.size > 0).length >= 1, null, { timeout: 10000 });


  const handleQuickSaveDownloadPromise = page.waitForEvent("download");


  await page.keyboard.press("ControlOrMeta+S");


  const handleQuickSaveDownload = await handleQuickSaveDownloadPromise;


  const handleQuickSavePath = await handleQuickSaveDownload.path();


  const handleQuickSaveHtml = await readFile(handleQuickSavePath, "utf8");


  saveAsDialog.handleQuickSave = await page.evaluate(() => ({
    writes: window.__webSheetSaveHandleWrites || [],
    status: document.querySelector("#statusText")?.textContent || "",
    workbookName: document.querySelector("#workbookName")?.textContent || "",
  }));


  saveAsDialog.handleQuickSave.writeCount = saveAsDialog.handleQuickSave.writes.filter((entry) => entry.size > 0).length;


  saveAsDialog.handleQuickSave.saveFileName = handleQuickSaveDownload.suggestedFilename();


  saveAsDialog.handleQuickSave.saveIsHtml = looksLikeStandaloneWorkbookHtml(handleQuickSaveHtml);


  saveAsDialog.handleQuickSave.saveHasWorkbook =
    handleQuickSaveHtml.includes("ZJ_スキルシート_20260517") ||
    handleQuickSaveHtml.includes("-picked");


  const expectedReloadWorkbookName = saveAsDialog.handleQuickSave?.workbookName || "ZJ_スキルシート_20260517.xlsx";



  const standaloneBlockedRequests = [];


  const standaloneLocalSaveRequests = [];


  const standalonePage = await context.newPage();


  standalonePage.on("console", (message) => {
    messages.push(`standalone ${message.type()}: ${message.text()}`);
  });


  standalonePage.on("pageerror", (error) => {
    messages.push(`standalone pageerror: ${error.stack || error.message}`);
  });


  await standalonePage.route("**/*", async (route) => {
    const url = route.request().url();
    if (/^http:\/\/(?:127\.0\.0\.1|localhost):\d+\/__websheet\/local-save/i.test(url)) {
      const postData = route.request().postData() || "";
      let payload = {};
      try {
        payload = JSON.parse(postData);
      } catch {
        payload = {};
      }
      standaloneLocalSaveRequests.push({
        url,
        hasWorkbookModel: String(payload.html || "").includes('id="websheet-model"'),
        path: payload.path || "",
      });
      await route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ ok: true }) });
      return;
    }
    if (/^https?:/i.test(url)) {
      standaloneBlockedRequests.push(url);
      await route.abort();
      return;
    }
    await route.continue();
  });


  await standalonePage.goto(pathToFileURL(exportedHtmlFilePath).href, { waitUntil: "load" });


  try {
    await standalonePage.waitForFunction(() => document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim() === "CHT", null, { timeout: 30000 });
  } catch (error) {
    const debug = await standalonePage.evaluate(() => ({
      title: document.title,
      bodyText: document.body?.innerText?.slice(0, 500) || "",
      tabText: document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim() || "",
      statusText: document.querySelector("#statusText")?.textContent?.trim() || "",
      scriptCount: document.scripts.length,
      runtimeStatus: document.querySelector("[data-runtime-status]")?.textContent?.trim() || "",
    })).catch((evaluateError) => ({ evaluateError: evaluateError.message || String(evaluateError) }));
    throw new Error(`Standalone exported HTML did not load workbook tabs: ${JSON.stringify({ exportedHtmlFilePath, debug })}`);
  }


  await standalonePage.waitForSelector("#sheetHost .sheet-cell[data-row][data-col]", { timeout: 10000 });


  const standaloneInitialStatus = await standalonePage.evaluate(() => document.querySelector("#statusText")?.textContent || "");


  await standalonePage.evaluate(() => {
    window.__standaloneQuickSavePickerCalls = [];
    window.__standaloneQuickSaveWrites = [];
    Object.defineProperty(window, "showSaveFilePicker", {
      configurable: true,
      value: async (options = {}) => {
        window.__standaloneQuickSavePickerCalls.push({
          suggestedName: String(options.suggestedName || ""),
          typeCount: Array.isArray(options.types) ? options.types.length : 0,
        });
        return {
          name: String(options.suggestedName || "websheet.html"),
          async createWritable() {
            return {
              async write(chunk) {
                window.__standaloneQuickSaveWrites.push({
                  size: Number(chunk?.size ?? chunk?.byteLength ?? 0),
                  type: String(chunk?.type || ""),
                });
              },
              async close() {
                window.__standaloneQuickSaveWrites.push({ closed: true });
              },
            };
          },
        };
      },
    });
	  });


	  await standalonePage.keyboard.press("Control+S");


	  await standalonePage.waitForFunction(() => (document.querySelector("#statusText")?.textContent || "").includes("上書き保存"), null, { timeout: 10000 });


	  const standaloneQuickSave = await standalonePage.evaluate(() => ({
	    pickerCalls: window.__standaloneQuickSavePickerCalls || [],
	    writes: window.__standaloneQuickSaveWrites || [],
	    status: document.querySelector("#statusText")?.textContent || "",
	    workbookName: document.querySelector("#workbookName")?.textContent || "",
	  }));


  standaloneQuickSave.localSaveRequests = standaloneLocalSaveRequests;


  await standalonePage.locator("#sheetHost .sheet-cell[data-row][data-col]").first().click({ button: "right" });


  await standalonePage.waitForSelector("#cellContextMenu", { timeout: 5000 });


  const standaloneCellMenuVisible = await standalonePage.locator("#cellContextMenu").isVisible();


  await standalonePage.keyboard.press("Escape");


  const standaloneLiveChart = await standalonePage.evaluate(() => {
    const hashText = (text) => {
      let hash = 0;
      for (let index = 0; index < text.length; index += 1) {
        hash = ((hash << 5) - hash + text.charCodeAt(index)) | 0;
      }
      return String(hash);
    };
    const sources = [
      ...[...document.querySelectorAll(".sheet-image, .sheet-image-link .sheet-image")].map((image) => image.getAttribute("src") || ""),
      ...((window.__WEBSHEET_STANDALONE_MODEL__?.sheets || [])
        .flatMap((sheet) => sheet.images || [])
        .filter((image) => image?.chart?.series?.length)
        .map((image) => image.src || "")),
    ].map((source) => decodeURIComponent(source || ""));
    const source = sources.find((candidate) => candidate.includes("上期")) || "";
    const modelCharts = (window.__WEBSHEET_STANDALONE_MODEL__?.sheets || [])
      .flatMap((sheet, sheetIndex) => (sheet.images || [])
        .filter((image) => image?.chart?.series?.length)
        .map((image) => ({
          sheetIndex,
          name: image.name || "",
          labels: image.chart?.labels || [],
          labelsRef: image.chart?.labelsRef || "",
          valuesRef: image.chart?.series?.[0]?.valuesRef || "",
          sourceHash: hashText(decodeURIComponent(image.src || "")),
        })));
    return {
      initialHasChart: Boolean(source),
      initialIncludesLabel: source.includes("上期"),
      initialSourceHash: hashText(source),
      imageCount: document.querySelectorAll(".sheet-image, .sheet-image-link .sheet-image").length,
      chartSourceCount: sources.filter((candidate) => candidate.includes("<svg") || candidate.includes("data:image/svg+xml")).length,
      modelCharts,
    };
  });


  const standaloneCells = standalonePage.locator("#sheetHost .sheet-cell[data-row][data-col]");


  await standaloneCells.nth(0).click();


  await standalonePage.keyboard.down("Control");


  await standaloneCells.nth(1).click();


  await standalonePage.keyboard.up("Control");


  await standalonePage.waitForFunction(() => document.querySelectorAll(".sheet-selection-frame").length >= 2 || document.querySelector(".sheet-selection-frame.is-multi-range"), null, { timeout: 5000 });


  await standaloneCells.nth(0).click();


  await standalonePage.evaluate(() => {
    window.__standaloneClipboardApiCalls = { available: Boolean(navigator.clipboard), writeText: 0, readText: 0 };
    const clipboard = navigator.clipboard;
    if (!clipboard) return;
    const originalWriteText = clipboard.writeText?.bind(clipboard);
    const originalReadText = clipboard.readText?.bind(clipboard);
    try {
      Object.defineProperty(clipboard, "writeText", {
        configurable: true,
        value: async (...args) => {
          window.__standaloneClipboardApiCalls.writeText += 1;
          return originalWriteText ? originalWriteText(...args) : undefined;
        },
      });
    } catch {}
    try {
      Object.defineProperty(clipboard, "readText", {
        configurable: true,
        value: async (...args) => {
          window.__standaloneClipboardApiCalls.readText += 1;
          return originalReadText ? originalReadText(...args) : "";
        },
      });
    } catch {}
	  });


	  const standaloneCopyEvent = await standalonePage.evaluate(() => {
	    let setDataCalls = 0;
	    const event = new Event("copy", { bubbles: true, cancelable: true });
	    Object.defineProperty(event, "clipboardData", {
	      configurable: true,
	      value: {
	        setData: () => {
	          setDataCalls += 1;
	        },
	      },
	    });
	    document.dispatchEvent(event);
	    return {
	      defaultPrevented: event.defaultPrevented,
	      setDataCalls,
	    };
	  });


	  await standalonePage.click('[data-ribbon-command="copy"]');


	  await standalonePage.keyboard.press("Escape");


	  await standalonePage.click('[data-ribbon-command="paste"]');


	  const standaloneToolbarClipboard = await standalonePage.evaluate(() => ({
	    calls: { ...(window.__standaloneClipboardApiCalls || {}) },
	    status: document.querySelector("#statusText")?.textContent || "",
	  }));


	  await standalonePage.keyboard.press("Control+C");


	  await standaloneCells.nth(1).click();


	  await standalonePage.keyboard.press("Control+V");


	  await standalonePage.waitForTimeout(220);


	  const standaloneClipboard = await standalonePage.evaluate(() => ({
	    calls: window.__standaloneClipboardApiCalls,
	    status: document.querySelector("#statusText")?.textContent || "",
	  }));


	  standaloneClipboard.toolbar = standaloneToolbarClipboard;


	  standaloneClipboard.copyEvent = standaloneCopyEvent;


  await standaloneCells.nth(0).click();


  await standalonePage.keyboard.down("Control");


  await standaloneCells.nth(1).click();


  await standalonePage.keyboard.up("Control");


  await standalonePage.waitForFunction(() => document.querySelectorAll(".sheet-selection-frame").length >= 2 || document.querySelector(".sheet-selection-frame.is-multi-range"), null, { timeout: 5000 });


  await standalonePage.click('[data-ribbon-tab="view"]');


  await standalonePage.waitForSelector('[data-ribbon-command="view-zoom"]', { timeout: 5000 });


  await standalonePage.click('[data-ribbon-command="view-zoom"]');


  await standalonePage.waitForSelector("#zoomDialog", { timeout: 5000 });


  const standaloneZoomDialog = await standalonePage.evaluate(() => ({
    title: document.querySelector(".zoom-dialog-title")?.textContent || "",
    selected: document.querySelector('#zoomDialog input[name="zoomDialogOption"]:checked')?.value || "",
    accessKeys: [...document.querySelectorAll("#zoomDialog [data-access-key]")].map((item) => item.dataset.accessKey || ""),
    accessKeysUnderlined: [...document.querySelectorAll("#zoomDialog .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  await standalonePage.keyboard.press("Alt+7");


  await standalonePage.waitForFunction(() => document.querySelector('#zoomDialog input[name="zoomDialogOption"][value="75"]')?.checked === true, null, { timeout: 5000 });


  await standalonePage.locator("#zoomDialog").press("Enter");


  await standalonePage.waitForFunction(() => {
    const sheetZoom = window.__webSheetDevSheetModel?.("CHT")?.zoomScale ??
      window.__WEBSHEET_STANDALONE_MODEL__?.sheets?.[0]?.zoomScale;
    return sheetZoom === 75;
  }, null, { timeout: 5000 });


  const standaloneZoomApplied = await standalonePage.evaluate(() => ({
    zoom: window.__webSheetDevSheetModel?.("CHT")?.zoomScale ??
      window.__WEBSHEET_STANDALONE_MODEL__?.sheets?.[0]?.zoomScale,
    label: document.querySelector("#zoomLabel")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await standalonePage.click('[data-ribbon-command="view-100"]');


  await standalonePage.waitForFunction(() => {
    const sheetZoom = window.__webSheetDevSheetModel?.("CHT")?.zoomScale ??
      window.__WEBSHEET_STANDALONE_MODEL__?.sheets?.[0]?.zoomScale;
    return sheetZoom === 100;
  }, null, { timeout: 5000 });


  await standalonePage.click('[data-ribbon-tab="home"]');


  const standaloneOffline = await standalonePage.evaluate((initialStatus) => ({
    firstTab: document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim(),
    cellCount: document.querySelectorAll("#sheetHost .sheet-cell").length,
    hasRibbon: Boolean(document.querySelector(".ribbon")),
    hasHomePanel: Boolean(document.querySelector('[data-ribbon-panel="home"].is-active')),
    selectionFrameCount: document.querySelectorAll(".sheet-selection-frame").length,
    hasMultiRangeFrame: Boolean(document.querySelector(".sheet-selection-frame.is-multi-range")),
    standaloneModelLoaded: Boolean(window.__WEBSHEET_STANDALONE_MODEL__),
    runtimeSource: window.__WEBSHEET_RUNTIME_SOURCE__,
    loadingStatusCleared: /HTML ワークブックを読み込みました|HTML 工作簿已加载/.test(initialStatus),
    status: initialStatus,
    cellMenuVisible: window.__standaloneCellMenuVisible || false,
  }), standaloneInitialStatus);


  standaloneOffline.cellMenuVisible = standaloneCellMenuVisible;


  standaloneOffline.quickSave = standaloneQuickSave;


  standaloneOffline.clipboard = standaloneClipboard;


  standaloneOffline.zoomDialog = { dialog: standaloneZoomDialog, applied: standaloneZoomApplied };


  await standalonePage.locator('#sheetHost .sheet-cell[data-row="2"][data-col="1"]').click();


  await standalonePage.waitForFunction(() =>
    document.querySelector("#selectedCellLabel")?.textContent?.trim() === "A2" &&
    document.querySelector("#formulaInput") &&
    !document.querySelector("#formulaInput").disabled,
  null, { timeout: 5000 });


  await standalonePage.locator("#formulaInput").fill("更新上期");


  await standalonePage.keyboard.press("Enter");


  await standalonePage.waitForFunction(() => {
    const model = window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "CHT", 2, 1);
    const domText = document.querySelector('#sheetHost .sheet-cell[data-row="2"][data-col="1"]')?.textContent || "";
    return model?.raw === "更新上期" || model?.display === "更新上期" || domText.includes("更新上期");
  }, null, { timeout: 10000 });


  await standalonePage.waitForFunction(() => {
    const sources = [
      ...[...document.querySelectorAll(".sheet-image, .sheet-image-link .sheet-image")].map((image) => image.getAttribute("src") || ""),
      ...((window.__WEBSHEET_STANDALONE_MODEL__?.sheets || [])
        .flatMap((sheet) => sheet.images || [])
        .filter((image) => image?.chart?.series?.length)
        .map((image) => image.src || "")),
    ].map((source) => decodeURIComponent(source || ""));
    return sources.some((source) => source.includes("更新上期"));
  }, null, { timeout: 5000 });


  const standaloneLiveChartAfterLabelHash = await standalonePage.evaluate(() => {
    const hashText = (text) => {
      let hash = 0;
      for (let index = 0; index < text.length; index += 1) {
        hash = ((hash << 5) - hash + text.charCodeAt(index)) | 0;
      }
      return String(hash);
    };
    const sources = [
      ...[...document.querySelectorAll(".sheet-image, .sheet-image-link .sheet-image")].map((image) => image.getAttribute("src") || ""),
      ...((window.__WEBSHEET_STANDALONE_MODEL__?.sheets || [])
        .flatMap((sheet) => sheet.images || [])
        .filter((image) => image?.chart?.series?.length)
        .map((image) => image.src || "")),
    ].map((source) => decodeURIComponent(source || ""));
    return hashText(sources.find((source) => source.includes("更新上期")) || "");
  });


  await standalonePage.locator('#sheetHost .sheet-cell[data-row="2"][data-col="2"]').click();


  await standalonePage.waitForFunction(() =>
    document.querySelector("#selectedCellLabel")?.textContent?.trim() === "B2" &&
    document.querySelector("#formulaInput") &&
    !document.querySelector("#formulaInput").disabled,
  null, { timeout: 5000 });


  await standalonePage.locator("#formulaInput").fill("333");


  await standalonePage.keyboard.press("Enter");


  await standalonePage.waitForFunction(() =>
    document.querySelector('#sheetHost .sheet-cell[data-row="2"][data-col="2"]')?.textContent?.includes("333"),
  null, { timeout: 5000 });


  await standalonePage.waitForFunction((previousHash) => {
    const hashText = (text) => {
      let hash = 0;
      for (let index = 0; index < text.length; index += 1) {
        hash = ((hash << 5) - hash + text.charCodeAt(index)) | 0;
      }
      return String(hash);
    };
    const sources = [
      ...[...document.querySelectorAll(".sheet-image, .sheet-image-link .sheet-image")].map((image) => image.getAttribute("src") || ""),
      ...((window.__WEBSHEET_STANDALONE_MODEL__?.sheets || [])
        .flatMap((sheet) => sheet.images || [])
        .filter((image) => image?.chart?.series?.length)
        .map((image) => image.src || "")),
    ].map((source) => decodeURIComponent(source || ""));
    const source = sources.find((candidate) => candidate.includes("更新上期")) || "";
    return Boolean(source) && hashText(source) !== previousHash;
  }, standaloneLiveChartAfterLabelHash, { timeout: 5000 });


  standaloneLiveChart.updated = await standalonePage.evaluate(() => {
    const hashText = (text) => {
      let hash = 0;
      for (let index = 0; index < text.length; index += 1) {
        hash = ((hash << 5) - hash + text.charCodeAt(index)) | 0;
      }
      return String(hash);
    };
    const sources = [
      ...[...document.querySelectorAll(".sheet-image, .sheet-image-link .sheet-image")].map((image) => image.getAttribute("src") || ""),
      ...((window.__WEBSHEET_STANDALONE_MODEL__?.sheets || [])
        .flatMap((sheet) => sheet.images || [])
        .filter((image) => image?.chart?.series?.length)
        .map((image) => image.src || "")),
    ].map((candidate) => decodeURIComponent(candidate || ""));
    const source = sources.find((candidate) => candidate.includes("更新上期")) || "";
    return {
      hasChart: Boolean(source),
      includesUpdatedLabel: source.includes("更新上期"),
      sourceHash: hashText(source),
    };
  });


  standaloneLiveChart.updated.valueSourceChanged = standaloneLiveChart.updated.sourceHash !== standaloneLiveChartAfterLabelHash;


  standaloneOffline.liveChart = standaloneLiveChart;


  await standalonePage.close();


  const remoteFallbackHtmlPath = path.join(path.dirname(exportedHtmlFilePath), "remote-runtime-fallback.html");


  const exportedLocalRuntimeBase = exportedHtml.match(/"localBaseUrl":"([^"]+)"/)?.[1] || runtimeBaseUrl;


  await writeFile(remoteFallbackHtmlPath, exportedHtml.replaceAll(exportedLocalRuntimeBase, `file:///tmp/websheet-runtime-missing/${STANDALONE_RUNTIME_CHANNEL}`), "utf8");


  const remoteFallbackPage = await context.newPage();


  remoteFallbackPage.on("console", (message) => {
    messages.push(`standalone fallback ${message.type()}: ${message.text()}`);
  });


  remoteFallbackPage.on("pageerror", (error) => {
    messages.push(`standalone fallback pageerror: ${error.stack || error.message}`);
  });


  await remoteFallbackPage.goto(pathToFileURL(remoteFallbackHtmlPath).href, { waitUntil: "load" });


  await remoteFallbackPage.waitForFunction(() => document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim() === "CHT", null, { timeout: 10000 });


  const standaloneRemoteFallback = await remoteFallbackPage.evaluate(() => ({
    firstTab: document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim(),
    cellCount: document.querySelectorAll("#sheetHost .sheet-cell").length,
    hasRibbon: Boolean(document.querySelector(".ribbon")),
    hasHomePanel: Boolean(document.querySelector('[data-ribbon-panel="home"].is-active')),
    standaloneModelLoaded: Boolean(window.__WEBSHEET_STANDALONE_MODEL__),
    runtimeSource: window.__WEBSHEET_RUNTIME_SOURCE__,
    loadingStatusCleared: /HTML ワークブックを読み込みました|HTML 工作簿已加载/.test(document.querySelector("#statusText")?.textContent || ""),
    status: document.querySelector("#statusText")?.textContent,
  }));


  await remoteFallbackPage.close();


  const exportedStandaloneModel = exportedStandaloneModelFromHtml(exportedHtml);


  const emptyModelHtmlPath = path.join(path.dirname(exportedHtmlFilePath), "websheet-empty-model.html");


  await writeFile(emptyModelHtmlPath, replaceStandaloneModelPayload(exportedHtml, {}), "utf8");


  const emptyModelPage = await context.newPage();


  const emptyModelBlockedRequests = [];


  emptyModelPage.on("console", (message) => {
    messages.push(`standalone empty model ${message.type()}: ${message.text()}`);
  });


  emptyModelPage.on("pageerror", (error) => {
    messages.push(`standalone empty model pageerror: ${error.stack || error.message}`);
  });


  await emptyModelPage.route("**/*", async (route) => {
    const url = route.request().url();
    if (/^https?:/i.test(url)) {
      emptyModelBlockedRequests.push(url);
      await route.abort();
      return;
    }
    await route.continue();
  });


  await emptyModelPage.goto(pathToFileURL(emptyModelHtmlPath).href, { waitUntil: "load" });


  await emptyModelPage.waitForSelector("#emptyState:not(.hidden)", { timeout: 10000 });


  const standaloneEmptyModel = await emptyModelPage.evaluate(() => ({
    emptyVisible: Boolean(document.querySelector("#emptyState:not(.hidden)")),
    workspaceHidden: Boolean(document.querySelector("#sheetWorkspace.hidden")),
    tabsBeforeLoad: document.querySelectorAll("#sheetTabs .sheet-tab").length,
    runtimeSource: window.__WEBSHEET_RUNTIME_SOURCE__,
    standaloneFlag: Boolean(window.__WEBSHEET_STANDALONE__),
    modelLoaded: Boolean(window.__WEBSHEET_STANDALONE_MODEL__),
    rawModelPresent: Boolean(window.__WEBSHEET_STANDALONE_MODEL_RAW__),
    status: document.querySelector("#statusText")?.textContent || "",
    alert: document.querySelector("#appAlert:not(.hidden)")?.textContent || "",
  }));


  await emptyModelPage.close();


  standaloneEmptyModel.blockedRequests = emptyModelBlockedRequests;



  const malformedModelHtmlPath = path.join(path.dirname(exportedHtmlFilePath), "websheet-malformed-model.html");


  await writeFile(malformedModelHtmlPath, replaceStandaloneModelScriptRaw(exportedHtml, '{"sheets":'), "utf8");


  const malformedModelPage = await context.newPage();


  const malformedModelBlockedRequests = [];


  malformedModelPage.on("console", (message) => {
    messages.push(`standalone malformed model ${message.type()}: ${message.text()}`);
  });


  malformedModelPage.on("pageerror", (error) => {
    messages.push(`standalone malformed model pageerror: ${error.stack || error.message}`);
  });


  await malformedModelPage.route("**/*", async (route) => {
    const url = route.request().url();
    if (/^https?:/i.test(url)) {
      malformedModelBlockedRequests.push(url);
      await route.abort();
      return;
    }
    await route.continue();
  });


  await malformedModelPage.goto(pathToFileURL(malformedModelHtmlPath).href, { waitUntil: "load" });


  await malformedModelPage.waitForSelector("#emptyState:not(.hidden)", { timeout: 10000 });


  const standaloneMalformedModel = await malformedModelPage.evaluate(() => ({
    emptyVisible: Boolean(document.querySelector("#emptyState:not(.hidden)")),
    workspaceHidden: Boolean(document.querySelector("#sheetWorkspace.hidden")),
    runtimeSource: window.__WEBSHEET_RUNTIME_SOURCE__,
    standaloneFlag: Boolean(window.__WEBSHEET_STANDALONE__),
    modelLoaded: Boolean(window.__WEBSHEET_STANDALONE_MODEL__),
    rawModelPresent: Boolean(window.__WEBSHEET_STANDALONE_MODEL_RAW__),
    modelError: String(window.__WEBSHEET_STANDALONE_MODEL_ERROR__ || ""),
    status: document.querySelector("#statusText")?.textContent || "",
    alert: document.querySelector("#appAlert:not(.hidden)")?.textContent || "",
  }));


  await malformedModelPage.close();


  standaloneMalformedModel.blockedRequests = malformedModelBlockedRequests;



  const wrappedModelHtmlPath = path.join(path.dirname(exportedHtmlFilePath), "websheet-wrapped-model.html");


  await writeFile(wrappedModelHtmlPath, replaceStandaloneModelPayload(exportedHtml, { workbook: exportedStandaloneModel }), "utf8");


  const wrappedModelPage = await context.newPage();


  const wrappedModelBlockedRequests = [];


  wrappedModelPage.on("console", (message) => {
    messages.push(`standalone wrapped model ${message.type()}: ${message.text()}`);
  });


  wrappedModelPage.on("pageerror", (error) => {
    messages.push(`standalone wrapped model pageerror: ${error.stack || error.message}`);
  });


  await wrappedModelPage.route("**/*", async (route) => {
    const url = route.request().url();
    if (/^https?:/i.test(url)) {
      wrappedModelBlockedRequests.push(url);
      await route.abort();
      return;
    }
    await route.continue();
  });


  await wrappedModelPage.goto(pathToFileURL(wrappedModelHtmlPath).href, { waitUntil: "load" });


  await wrappedModelPage.waitForFunction(() => document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim() === "CHT", null, { timeout: 10000 });


  const standaloneWrappedModel = await wrappedModelPage.evaluate(() => ({
    firstTab: document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim(),
    cellCount: document.querySelectorAll("#sheetHost .sheet-cell").length,
    hasRibbon: Boolean(document.querySelector(".ribbon")),
    hasHomePanel: Boolean(document.querySelector('[data-ribbon-panel="home"].is-active')),
    workspaceVisible: Boolean(document.querySelector("#sheetWorkspace:not(.hidden)")),
    runtimeSource: window.__WEBSHEET_RUNTIME_SOURCE__,
    standaloneFlag: Boolean(window.__WEBSHEET_STANDALONE__),
    modelLoaded: Boolean(window.__WEBSHEET_STANDALONE_MODEL__),
    rawModelPresent: Boolean(window.__WEBSHEET_STANDALONE_MODEL_RAW__),
    sourceName: window.__webSheetDevWorkbookState?.().sourceName || document.querySelector("#workbookName")?.textContent || "",
    status: document.querySelector("#statusText")?.textContent || "",
  }));


  await wrappedModelPage.close();


  standaloneWrappedModel.blockedRequests = wrappedModelBlockedRequests;



  const launcherHtmlPath = path.join(path.dirname(exportedHtmlFilePath), "websheet-launcher-empty.html");


  await writeFile(
    launcherHtmlPath,
    exportedHtml
      .replace(/<title>[\s\S]*?<\/title>/, "<title>WebSheet Launcher</title>")
      .replace(WEBSHEET_MODEL_SCRIPT_RE, "")
      .replace(/\n\s*\n/, "\n"),
    "utf8",
  );


  const launcherBlockedRequests = [];


  const launcherPage = await context.newPage();


  launcherPage.on("console", (message) => {
    messages.push(`standalone launcher ${message.type()}: ${message.text()}`);
  });


  launcherPage.on("pageerror", (error) => {
    messages.push(`standalone launcher pageerror: ${error.stack || error.message}`);
  });


  await launcherPage.route("**/*", async (route) => {
    const url = route.request().url();
    if (/^https?:/i.test(url)) {
      launcherBlockedRequests.push(url);
      await route.abort();
      return;
    }
    await route.continue();
  });


  await launcherPage.goto(pathToFileURL(launcherHtmlPath).href, { waitUntil: "load" });


  await launcherPage.waitForSelector("#emptyState:not(.hidden)", { timeout: 10000 });


  await launcherPage.waitForSelector("#fileInput", { timeout: 10000 });


  const launcherBeforeLoad = await launcherPage.evaluate(() => ({
    hasRibbon: Boolean(document.querySelector(".ribbon")),
    emptyVisible: Boolean(document.querySelector("#emptyState:not(.hidden)")),
    workspaceHidden: Boolean(document.querySelector("#sheetWorkspace.hidden")),
    tabsBeforeLoad: document.querySelectorAll("#sheetTabs .sheet-tab").length,
    runtimeSource: window.__WEBSHEET_RUNTIME_SOURCE__,
    standaloneFlag: Boolean(window.__WEBSHEET_STANDALONE__),
    modelLoaded: Boolean(window.__WEBSHEET_STANDALONE_MODEL__),
    title: document.title,
  }));


  await launcherPage.setInputFiles("#fileInput", workbookPath);


  await launcherPage.waitForFunction(() => document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim() === "CHT", null, { timeout: 20000 });


  const standaloneLauncher = await launcherPage.evaluate(() => ({
    beforeLoad: window.__launcherBeforeLoad || null,
    firstTab: document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim(),
    workspaceVisible: Boolean(document.querySelector("#sheetWorkspace:not(.hidden)")),
    workbookName: document.querySelector("#workbookName")?.textContent,
    status: document.querySelector("#statusText")?.textContent,
    cellCount: document.querySelectorAll("#sheetHost .sheet-cell").length,
    runtimeSource: window.__WEBSHEET_RUNTIME_SOURCE__,
    standaloneFlag: Boolean(window.__WEBSHEET_STANDALONE__),
    modelLoaded: Boolean(window.__WEBSHEET_STANDALONE_MODEL__),
  }));


  standaloneLauncher.beforeLoad = launcherBeforeLoad;


  await launcherPage.click('[data-ribbon-tab="file"]');


  await launcherPage.waitForFunction(() => !document.querySelector("#fileSideMenu")?.classList.contains("hidden"), null, { timeout: 10000 });


  await launcherPage.click('#fileSideMenu [data-ribbon-command="save-as"]');


  await launcherPage.waitForSelector('[data-data-dialog="save-as"]', { timeout: 10000 });


  standaloneLauncher.saveAsDefault = await launcherPage.evaluate(() => ({
    fileName: document.querySelector("[data-data-field='saveAsFileName']")?.value || "",
    type: document.querySelector("[data-data-field='saveAsType']")?.value || "",
    activeType: document.querySelector("[data-save-as-file-row].is-active")?.dataset.saveAsFileType || "",
    activeFileName: document.querySelector("[data-save-as-file-row].is-active")?.dataset.saveAsFileName || "",
    note: document.querySelector("[data-save-as-note]")?.textContent || "",
  }));


  await launcherPage.click('[data-data-dialog-action="cancel"]');


  await launcherPage.waitForSelector('[data-data-dialog="save-as"]', { state: "detached", timeout: 5000 });


  const launcherQuickSavePromise = launcherPage.waitForEvent("download");


  await launcherPage.keyboard.press("Control+S");


  const launcherQuickSaveDownload = await launcherQuickSavePromise;


  const launcherQuickSavePath = await launcherQuickSaveDownload.path();


  const launcherQuickSaveBytes = await readFile(launcherQuickSavePath);


  standaloneLauncher.quickSave = {
    fileName: launcherQuickSaveDownload.suggestedFilename(),
    zipHeader: launcherQuickSaveBytes.subarray(0, 4).toString("hex"),
    isHtml: looksLikeStandaloneWorkbookHtml(launcherQuickSaveBytes.toString("utf8", 0, Math.min(2048, launcherQuickSaveBytes.length))),
    status: await launcherPage.evaluate(() => document.querySelector("#statusText")?.textContent || ""),
    workbookName: await launcherPage.evaluate(() => document.querySelector("#workbookName")?.textContent || ""),
  };


  await launcherPage.close();



  await page.evaluate(async () => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    window.__webSheetDevSetActiveSheet?.("CHT");
    await frame();
    const clickCommand = async (command) => {
      document.querySelector(`[data-ribbon-command="${command}"]`)?.click();
      await frame();
    };
    const sheetSnapshot = () => window.__webSheetDevSheetModel?.("CHT") || {};
    if ((sheetSnapshot().viewMode || "normal") !== "normal") await clickCommand("view-normal");
    if (sheetSnapshot().showGridLines !== false) await clickCommand("view-gridlines");
    await frame();
  });


  await page.waitForFunction(
    () =>
      window.__webSheetDevWorkbookState?.().activeSheetName === "CHT" &&
      (window.__webSheetDevSheetModel?.("CHT")?.viewMode || "normal") === "normal" &&
      window.__webSheetDevSheetModel?.("CHT")?.showGridLines === false &&
      document.querySelector(".sheet-grid")?.classList.contains("hide-gridlines"),
    null,
    { timeout: 10000 },
  );


  let ar2LazyDebug = null;


  for (const offset of [0, -480, 480, -960, 960, -1440, 1440]) {
    ar2LazyDebug = await page.evaluate(async (scrollOffset) => {
      const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const host = document.querySelector("#sheetHost");
      const sheet = window.__webSheetDevSheetModel?.("CHT");
      const targetCol = 44;
      const rowHeaderWidth = Math.round(document.querySelector(".row-header")?.getBoundingClientRect().width || 28);
      const targetLeft =
        rowHeaderWidth +
        (sheet?.columns || [])
          .slice(0, targetCol - 1)
          .reduce((total, column) => total + (column?.hidden ? 0 : Number(column?.width || sheet?.defaultColumnWidth || 0)), 0);
      host.scrollTop = 0;
      host.scrollLeft = Math.max(0, targetLeft - Math.max(80, host.clientWidth * 0.38) + scrollOffset);
      host.dispatchEvent(new Event("scroll"));
      await frame();
      await new Promise((resolve) => setTimeout(resolve, 80));
      const headers = [...document.querySelectorAll(".column-header[data-col]")].map((header) => header.textContent.trim());
      return {
        scrollOffset,
        targetLeft,
        activeSheetName: window.__webSheetDevWorkbookState?.().activeSheetName || "",
        scrollLeft: host.scrollLeft,
        scrollWidth: host.scrollWidth,
        clientWidth: host.clientWidth,
        sheetColCount: window.__webSheetDevSheetModel?.("CHT")?.colCount || 0,
        hasAR2: Boolean(document.querySelector('.sheet-cell[data-address="AR2"]')),
        hasAQ2: Boolean(document.querySelector('.sheet-cell[data-address="AQ2"]')),
        hasAS2: Boolean(document.querySelector('.sheet-cell[data-address="AS2"]')),
        renderedHeadersTail: headers.slice(-12),
      };
    }, offset);
    if (ar2LazyDebug.hasAR2) break;
  }


  if (!ar2LazyDebug?.hasAR2) {
    throw new Error(`AR2 was not rendered after horizontal lazy scroll: ${JSON.stringify(ar2LazyDebug)}`);
  }


  const ar2Gridline = await page.evaluate(() => {
    const inspect = (address) => {
      const cell = document.querySelector(`.sheet-cell[data-address="${address}"]`);
      const style = cell ? getComputedStyle(cell) : null;
      return {
        className: cell?.className,
        fillColor: style?.getPropertyValue("--cell-fill-color").trim().toLowerCase(),
        backgroundColor: style?.backgroundColor,
        borderRightColor: style?.borderRightColor,
        borderBottomColor: style?.borderBottomColor,
      };
    };
    return {
      gridClassName: document.querySelector(".sheet-grid")?.className,
      ar2: inspect("AR2"),
      aq2: inspect("AQ2"),
      ar1: inspect("AR1"),
    };
  });



  await page.click('.sheet-tab[data-sheet-index="1"]');


  await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent === "AI製品紹介", null, { timeout: 10000 });


  await page.waitForFunction(
    () => {
      const images = [...document.querySelectorAll(".sheet-image")];
      return images.length >= 1 && images.every((image) => image.complete && image.naturalWidth > 0);
    },
    null,
    { timeout: 10000 },
  );


  const collectImageAndFont = async () =>
    page.evaluate(() => {
    const gridRect = document.querySelector(".sheet-grid")?.getBoundingClientRect();
    const rowHeaderWidth = Math.round(document.querySelector(".row-header")?.getBoundingClientRect().width || 0);
    const a10Frame = document.querySelector('.sheet-cell[data-address="A10"]')?.getBoundingClientRect();
    const images = [...document.querySelectorAll(".sheet-image")].map((image) => {
      const rect = image.getBoundingClientRect();
      return {
        id: image.dataset.imageId || image.parentElement?.dataset.imageId || "",
        src: image.getAttribute("src")?.slice(0, 32),
        naturalWidth: image.naturalWidth,
        naturalHeight: image.naturalHeight,
        width: Math.round(rect.width),
        height: Math.round(rect.height),
        left: Math.round(rect.left - (gridRect?.left || 0)),
        top: Math.round(rect.top - (gridRect?.top || 0)),
      };
    });
    const frame = a10Frame
      ? {
          left: Math.round(a10Frame.left - (gridRect?.left || 0)),
          right: Math.round(a10Frame.right - (gridRect?.left || 0)),
          width: Math.round(a10Frame.width),
        }
      : null;
    const a1 = document.querySelector('.sheet-cell[data-address="A1"]');
    const a2 = document.querySelector('.sheet-cell[data-address="A2"]');
    const b6 = document.querySelector('.sheet-cell[data-address="B6"]');
    const a2Style = a2 ? getComputedStyle(a2) : null;
    return {
      imageCount: images.length,
      images,
      frame,
      rowHeaderWidth,
      a1FontFamily: a1 ? getComputedStyle(a1).fontFamily : "",
      b6FontFamily: b6 ? getComputedStyle(b6).fontFamily : "",
      a2WhiteSpace: a2Style?.whiteSpace || "",
      a2OverflowWrap: a2Style?.overflowWrap || "",
    };
  });


  const imageAndFontTop = await collectImageAndFont();


  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    if (host) host.scrollTop = 850;
  });


  await page.waitForFunction(
    () => {
      const host = document.querySelector("#sheetHost");
      const images = [...document.querySelectorAll(".sheet-image")];
      return (
        (host?.scrollTop || 0) >= 800 &&
        images.length >= 1 &&
        images.every((image) => image.complete && image.naturalWidth > 0) &&
        images.some((image) => {
          const rect = image.getBoundingClientRect();
          return rect.bottom > 0 && rect.top < window.innerHeight;
        })
      );
    },
    null,
    { timeout: 10000 },
  );


  const imageAndFontScrolled = await collectImageAndFont();


  const imageMap = new Map([...imageAndFontTop.images, ...imageAndFontScrolled.images].map((image) => [image.id || `${image.left}:${image.top}`, image]));


  const imageAndFont = {
    ...imageAndFontTop,
    imageCount: imageMap.size,
    images: [...imageMap.values()],
    renderedAtTop: imageAndFontTop.images.length,
    renderedAfterScroll: imageAndFontScrolled.images.length,
  };


  const imageSelection = await page.evaluate(() => {
    const image = [...document.querySelectorAll(".sheet-image")].find((node) => {
      const rect = node.getBoundingClientRect();
      return rect.top >= 0 && rect.top < window.innerHeight && rect.left >= 0 && rect.left < window.innerWidth;
    });
    image?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, clientX: image.getBoundingClientRect().left + 8, clientY: image.getBoundingClientRect().top + 8 }));
    return {
      selectedImages: document.querySelectorAll(".sheet-image.is-selected").length,
      selectedLabel: document.querySelector("#selectedCellLabel")?.textContent,
      formulaDisabled: document.querySelector("#formulaInput")?.disabled,
    };
  });



  await page.click('.sheet-tab[data-sheet-index="0"]');


  await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent === "CHT", null, { timeout: 10000 });


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B14"]'), null, { timeout: 10000 });



  const footerLayout = await page.evaluate(async () => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const rectFor = (selector) => {
      const rect = document.querySelector(selector)?.getBoundingClientRect();
      return rect
        ? {
            left: rect.left,
            top: rect.top,
            right: rect.right,
            bottom: rect.bottom,
            width: rect.width,
            height: rect.height,
          }
        : null;
    };
    const footer = rectFor("#sheetFooter");
    const tabs = rectFor(".sheet-tabs-pane");
    const right = rectFor(".sheet-footer-right");
    const scrollbar = rectFor("#horizontalScrollbar");
    const toolbar = rectFor(".sheet-view-toolbar");
    const b14Before = document.querySelector('.sheet-cell[data-address="B14"]')?.getBoundingClientRect().width || 0;
    const zoomSlider = document.querySelector("#zoomSlider");
    const initialZoomLabel = document.querySelector("#zoomLabel")?.textContent;
    zoomSlider.value = "100";
    zoomSlider.dispatchEvent(new Event("input", { bubbles: true }));
    await frame();
    const b14At100 = document.querySelector('.sheet-cell[data-address="B14"]')?.getBoundingClientRect().width || 0;
    const zoom100Label = document.querySelector("#zoomLabel")?.textContent;
    zoomSlider.value = "118";
    zoomSlider.dispatchEvent(new Event("input", { bubbles: true }));
    await frame();
    const wheelGrowDefaultPrevented = !document.querySelector("#sheetHost").dispatchEvent(new WheelEvent("wheel", { bubbles: true, cancelable: true, ctrlKey: true, deltaY: -100 }));
    await frame();
    const zoomAfterWheelGrow = document.querySelector("#zoomLabel")?.textContent;
    const b14AfterWheelGrow = document.querySelector('.sheet-cell[data-address="B14"]')?.getBoundingClientRect().width || 0;
    const wheelShrinkDefaultPrevented = !document.querySelector("#sheetHost").dispatchEvent(new WheelEvent("wheel", { bubbles: true, cancelable: true, ctrlKey: true, deltaY: 100 }));
    await frame();
    const zoomAfterWheelShrink = document.querySelector("#zoomLabel")?.textContent;
    const b14AfterWheelShrink = document.querySelector('.sheet-cell[data-address="B14"]')?.getBoundingClientRect().width || 0;
    document.querySelector('[data-sheet-view="page-layout"]')?.click();
    const activeAfterPageLayout = document.querySelector(".sheet-view-button.is-active")?.getAttribute("aria-label");
    document.querySelector('[data-sheet-view="normal"]')?.click();
    return {
      footer,
      tabs,
      right,
      scrollbar,
      toolbar,
      ratio: tabs && right ? tabs.width / (tabs.width + right.width) : 0,
      scrollbarAboveToolbar: Boolean(scrollbar && toolbar && scrollbar.bottom <= toolbar.top + 1),
      viewLabels: [...document.querySelectorAll("[data-sheet-view]")].map((button) => button.getAttribute("aria-label")),
      visibleViewText: [...document.querySelectorAll("[data-sheet-view]")].map((button) => button.textContent.trim()).join(""),
      markCount: document.querySelectorAll(".sheet-view-mark").length,
      activeAfterPageLayout,
      activeAfterNormal: document.querySelector(".sheet-view-button.is-active")?.getAttribute("aria-label"),
      initialZoomLabel,
      zoom100Label,
      wheelGrowDefaultPrevented,
      wheelShrinkDefaultPrevented,
      zoomAfterWheelGrow,
      zoomAfterWheelShrink,
      restoredZoomLabel: document.querySelector("#zoomLabel")?.textContent,
      b14Before,
      b14At100,
      b14AfterWheelGrow,
      b14AfterWheelShrink,
      b14Restored: document.querySelector('.sheet-cell[data-address="B14"]')?.getBoundingClientRect().width || 0,
    };
  });



  const realCtrlWheelZoom = {
    hasHost: false,
    growLabels: [],
    shrinkLabels: [],
    undoLabel: null,
    undoStatus: null,
    redoLabel: null,
    redoStatus: null,
  };


  const sheetHostBox = await page.locator("#sheetHost").boundingBox();


  if (sheetHostBox) {
    realCtrlWheelZoom.hasHost = true;
    await page.mouse.move(sheetHostBox.x + sheetHostBox.width / 2, sheetHostBox.y + sheetHostBox.height / 2);
    await page.keyboard.down("Control");
    try {
      for (let i = 0; i < 4; i += 1) {
        await page.mouse.wheel(0, -100);
        await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
        realCtrlWheelZoom.growLabels.push(await page.locator("#zoomLabel").textContent());
      }
      for (let i = 0; i < 4; i += 1) {
        await page.mouse.wheel(0, 100);
        await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
        realCtrlWheelZoom.shrinkLabels.push(await page.locator("#zoomLabel").textContent());
      }
    } finally {
      await page.keyboard.up("Control");
    }
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(() => document.querySelector("#zoomLabel")?.textContent === "128%", null, { timeout: 10000 });
    Object.assign(realCtrlWheelZoom, await page.evaluate(() => ({
      undoLabel: document.querySelector("#zoomLabel")?.textContent,
      undoStatus: document.querySelector("#statusText")?.textContent,
    })));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(() => document.querySelector("#zoomLabel")?.textContent === "118%", null, { timeout: 10000 });
    Object.assign(realCtrlWheelZoom, await page.evaluate(() => ({
      redoLabel: document.querySelector("#zoomLabel")?.textContent,
      redoStatus: document.querySelector("#statusText")?.textContent,
    })));
  }



  const zoomPopup = await page.evaluate(async () => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const openDialog = async () => {
      document.querySelector("#zoomLabel").click();
      await frame();
      return document.querySelector("#zoomDialog");
    };

    let dialog = await openDialog();
    const labelRect = document.querySelector("#zoomLabel").getBoundingClientRect();
    const dialogRect = dialog?.getBoundingClientRect();
    const initial = {
      visible: Boolean(dialog),
      title: dialog?.querySelector(".zoom-dialog-title")?.textContent,
      sectionTitle: dialog?.querySelector(".zoom-dialog-section-title")?.textContent,
      options: [...(dialog?.querySelectorAll(".zoom-dialog-option") || [])].map((option) => option.textContent.replace(/\s+/g, " ").trim()),
      accessKeys: [...(dialog?.querySelectorAll(".menu-access-key") || [])].map((item) => item.textContent),
      accessKeysUnderlined: [...(dialog?.querySelectorAll(".menu-access-key") || [])].every((item) =>
        getComputedStyle(item).textDecorationLine.includes("underline"),
      ),
      selectedOption: dialog?.querySelector('input[name="zoomDialogOption"]:checked')?.value,
      customValue: dialog?.querySelector("[data-zoom-custom-input]")?.value,
      popupAboveLabel: Boolean(dialogRect && dialogRect.bottom <= labelRect.top + 1),
    };

    dialog.querySelector('input[name="zoomDialogOption"][value="75"]').click();
    dialog.querySelector('[data-zoom-dialog-action="ok"]').click();
    await frame();
    const afterPreset = document.querySelector("#zoomLabel")?.textContent;
    const dialogClosedAfterPreset = !document.querySelector("#zoomDialog");

    dialog = await openDialog();
    dialog.querySelector('input[name="zoomDialogOption"][value="custom"]').click();
    const customInput = dialog.querySelector("[data-zoom-custom-input]");
    customInput.value = "118";
    customInput.dispatchEvent(new Event("input", { bubbles: true }));
    dialog.querySelector('[data-zoom-dialog-action="ok"]').click();
    await frame();
    const afterCustom = document.querySelector("#zoomLabel")?.textContent;

    dialog = await openDialog();
    dialog.querySelector('input[name="zoomDialogOption"][value="50"]').click();
    dialog.querySelector('[data-zoom-dialog-action="cancel"]').click();
    await frame();
    const afterCancel = document.querySelector("#zoomLabel")?.textContent;

    return {
      ...initial,
      afterPreset,
      dialogClosedAfterPreset,
      afterCustom,
      afterCancel,
      closedAfterCancel: !document.querySelector("#zoomDialog"),
    };
  });



  await page.evaluate(() => {
    const scroller = document.querySelector("#horizontalScrollbar");
    scroller.scrollLeft = 300;
    scroller.dispatchEvent(new Event("scroll"));
  });


  await page.waitForTimeout(100);



  const footerBeforeScroll = await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    const footer = document.querySelector("#sheetFooter");
    host.scrollTop = 0;
    const before = footer.getBoundingClientRect();
    host.scrollTop = 700;
    host.dispatchEvent(new Event("scroll"));
    const after = footer.getBoundingClientRect();
    return {
      viewportHeight: window.innerHeight,
      pageScrollY: window.scrollY,
      beforeBottom: before.bottom,
      afterBottom: after.bottom,
      beforeTop: before.top,
      afterTop: after.top,
      hostScrollTop: host.scrollTop,
    };
  });



  const tabScroll = await page.evaluate(async () => {
    const footer = document.querySelector("#sheetFooter");
    const tabs = document.querySelector("#sheetTabs");
    footer.style.setProperty("--sheet-tabs-width", "130px");
    tabs.scrollLeft = 0;
    document.querySelector("#sheetTabScrollRight").click();
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const afterRight = tabs.scrollLeft;
    document.querySelector("#sheetTabScrollLeft").click();
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    return {
      maxScrollLeft: tabs.scrollWidth - tabs.clientWidth,
      afterRight,
      afterLeft: tabs.scrollLeft,
    };
  });



  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="B14"]'), null, { timeout: 10000 });


  await page.locator('.sheet-cell[data-address="B14"]').click();


  const singleClickEdit = await page.evaluate(() => document.querySelector('.sheet-cell[data-address="B14"]')?.getAttribute("contenteditable"));


  await page.evaluate(() => {
    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
    if (!cell) return;
    cell.style.textAlign = "right";
    cell.style.justifyContent = "flex-end";
    cell.style.alignItems = "flex-end";
  });


  await page.locator('.sheet-cell[data-address="B14"]').dblclick({ position: { x: 8, y: 8 } });


  const doubleClickEdit = await page.evaluate(() => document.querySelector('.sheet-cell[data-address="B14"]')?.getAttribute("contenteditable"));


  await page.waitForTimeout(500);


  const editAfterDeferredRender = await page.evaluate(() => ({
    editable: document.querySelector('.sheet-cell[data-address="B14"]')?.getAttribute("contenteditable"),
    activeAddress: document.activeElement?.dataset?.address || "",
  }));


  const editAlignment = await page.evaluate(() => {
    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
    const style = cell ? getComputedStyle(cell) : null;
    return {
      editable: cell?.getAttribute("contenteditable"),
      display: style?.display || "",
      textAlign: style?.textAlign || "",
      justifyContent: style?.justifyContent || "",
      alignItems: style?.alignItems || "",
      paddingTopPx: Number.parseFloat(style?.paddingTop || "0") || 0,
    };
  });


  const editBox = await page.locator('.sheet-cell[data-address="B14"]').boundingBox();


  if (!editBox) throw new Error("B14 edit cell was not visible.");


  await page.mouse.click(editBox.x + Math.min(14, Math.max(4, editBox.width - 4)), editBox.y + Math.min(12, Math.max(4, editBox.height - 4)));


  await page.waitForTimeout(80);


  const editClickCaret = await page.evaluate(() => {
    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
    const selection = window.getSelection();
    return {
      editable: cell?.getAttribute("contenteditable"),
      activeAddress: document.activeElement?.dataset?.address || "",
      selectionInCell:
        Boolean(cell && selection?.anchorNode && cell.contains(selection.anchorNode)) &&
        Boolean(cell && selection?.focusNode && cell.contains(selection.focusNode)),
    };
  });


  await page.mouse.move(editBox.x + 4, editBox.y + Math.min(12, Math.max(4, editBox.height - 4)));


  await page.mouse.down();


  await page.mouse.move(editBox.x + Math.max(20, editBox.width - 4), editBox.y + Math.min(12, Math.max(4, editBox.height - 4)), { steps: 8 });


  await page.mouse.up();


  await page.waitForTimeout(120);


	  const editTextSelection = await page.evaluate(() => {
	    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
	    const selection = window.getSelection();
	    return {
	      editable: cell?.getAttribute("contenteditable"),
      selectedText: selection?.toString() || "",
      selectionInCell:
        Boolean(cell && selection?.anchorNode && cell.contains(selection.anchorNode)) &&
	        Boolean(cell && selection?.focusNode && cell.contains(selection.focusNode)),
	    };
	  });


		  const editNoEllipsis = await page.evaluate(() => {
		    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
		    if (!cell) return null;
		    cell.textContent = "A".repeat(200);
		    const style = getComputedStyle(cell);
	    return {
	      editable: cell.getAttribute("contenteditable"),
	      textOverflow: style.textOverflow,
	      overflowX: style.overflowX,
	      scrollsHorizontally: cell.scrollWidth > cell.clientWidth,
		      text: cell.textContent,
		    };
		  });


	  await page.evaluate(() => {
	    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
	    if (!cell) return;
	    cell.focus();
	    cell.textContent = "Line1";
	    const selection = window.getSelection();
	    const range = document.createRange();
	    range.selectNodeContents(cell);
	    range.collapse(false);
	    selection.removeAllRanges();
	    selection.addRange(range);
	  });


	  await page.keyboard.down("Alt");


	  await page.keyboard.press("Enter");


	  await page.keyboard.up("Alt");


	  await page.waitForTimeout(80);


	  const cellEditAltEnter = await page.evaluate(() => {
	    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
	    if (!cell) return null;
	    const style = getComputedStyle(cell);
	    return {
	      editable: cell.getAttribute("contenteditable"),
	      text: cell.textContent,
	      lineCount: (cell.textContent.match(/\n/g) || []).length + 1,
	      whiteSpace: style.whiteSpace,
	      activeAddress: document.activeElement?.dataset?.address || "",
	      label: document.querySelector("#selectedCellLabel")?.textContent || "",
	    };
	  });


	  const cellEditAltEnterDuringIme = await page.evaluate(async () => {
	    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
	    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
	    if (!cell) return null;
	    const compositionEvent = (type, data) => {
	      if (typeof CompositionEvent === "function") {
	        return new CompositionEvent(type, { bubbles: true, cancelable: false, data });
	      }
	      const event = new Event(type, { bubbles: true, cancelable: false });
	      Object.defineProperty(event, "data", { value: data });
	      return event;
	    };
	    cell.focus();
	    cell.textContent = "IME";
	    const selection = window.getSelection();
	    const range = document.createRange();
	    range.selectNodeContents(cell);
	    range.collapse(false);
	    selection.removeAllRanges();
	    selection.addRange(range);
	    cell.dispatchEvent(compositionEvent("compositionstart", "ime"));
	    const enterEvent = new KeyboardEvent("keydown", {
	      key: "Enter",
	      code: "Enter",
	      bubbles: true,
	      cancelable: true,
	      altKey: true,
	      isComposing: true,
	    });
	    Object.defineProperty(enterEvent, "keyCode", { get: () => 229 });
	    Object.defineProperty(enterEvent, "which", { get: () => 229 });
	    cell.dispatchEvent(enterEvent);
	    await frame();
	    const result = {
	      defaultPrevented: enterEvent.defaultPrevented,
	      editable: cell.getAttribute("contenteditable"),
	      text: cell.textContent,
	      activeAddress: document.activeElement?.dataset?.address || "",
	      activeElementInCell: document.activeElement === cell,
	      label: document.querySelector("#selectedCellLabel")?.textContent || "",
	    };
	    cell.dispatchEvent(compositionEvent("compositionend", "IME"));
	    await frame();
	    return result;
	  });


	  const editDragOutsideSelection = await page.evaluate(async () => {
	    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
	    const edit = document.querySelector('.sheet-cell[data-address="B14"]');
	    const target =
	      document.querySelector('.sheet-cell[data-address="D15"]') ||
	      [...document.querySelectorAll(".sheet-cell[data-address]")].find((cell) => cell !== edit && !cell.hasAttribute("contenteditable"));
	    if (!edit || !target) return null;
	    edit.focus();
	    const editRect = edit.getBoundingClientRect();
	    const targetRect = target.getBoundingClientRect();
	    const startX = editRect.left + Math.min(8, Math.max(2, editRect.width / 2));
	    const startY = editRect.top + Math.min(8, Math.max(2, editRect.height / 2));
	    const endX = targetRect.left + Math.min(8, Math.max(2, targetRect.width / 2));
	    const endY = targetRect.top + Math.min(8, Math.max(2, targetRect.height / 2));
	    const pointerId = 77;
	    edit.dispatchEvent(
	      new PointerEvent("pointerdown", { bubbles: true, button: 0, buttons: 1, pointerId, clientX: startX, clientY: startY }),
	    );
	    window.dispatchEvent(
	      new PointerEvent("pointermove", { bubbles: true, button: 0, buttons: 1, pointerId, clientX: endX, clientY: endY }),
	    );
	    window.dispatchEvent(
	      new PointerEvent("pointerup", { bubbles: true, button: 0, buttons: 0, pointerId, clientX: endX, clientY: endY }),
	    );
	    target.dispatchEvent(new MouseEvent("click", { bubbles: true, button: 0, clientX: endX, clientY: endY }));
	    await frame();
	    return {
	      targetAddress: target.dataset.address || "",
	      editable: edit.getAttribute("contenteditable"),
	      activeAddress: document.activeElement?.dataset?.address || "",
	      label: document.querySelector("#selectedCellLabel")?.textContent || "",
	    };
	  });


		  const cellEditImeEnter = await page.evaluate(async () => {
		    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
		    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
    if (!cell) return null;
    const compositionEvent = (type, data) => {
      if (typeof CompositionEvent === "function") {
        return new CompositionEvent(type, { bubbles: true, cancelable: false, data });
      }
      const event = new Event(type, { bubbles: true, cancelable: false });
      Object.defineProperty(event, "data", { value: data });
      return event;
    };
    cell.textContent = "変換中";
    cell.dispatchEvent(compositionEvent("compositionstart", "へんかん"));
    const enterEvent = new KeyboardEvent("keydown", {
      key: "Enter",
      code: "Enter",
      bubbles: true,
      cancelable: true,
      isComposing: true,
    });
    Object.defineProperty(enterEvent, "keyCode", { get: () => 229 });
    Object.defineProperty(enterEvent, "which", { get: () => 229 });
    cell.dispatchEvent(enterEvent);
    await frame();
    const during = {
      defaultPrevented: enterEvent.defaultPrevented,
      editable: cell.getAttribute("contenteditable"),
      activeAddress: document.activeElement?.dataset?.address || "",
      activeElementInCell: document.activeElement === cell,
      label: document.querySelector("#selectedCellLabel")?.textContent,
      text: cell.textContent,
    };
    cell.dispatchEvent(compositionEvent("compositionend", "変換中"));
    await frame();
    return {
      ...during,
      afterEndEditable: cell.getAttribute("contenteditable"),
      afterEndActiveAddress: document.activeElement?.dataset?.address || "",
    };
  });


  const editScrollbar = await page.evaluate(() => {
    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
    if (!cell) return null;
    const style = getComputedStyle(cell);
    const webkitScrollbar = getComputedStyle(cell, "::-webkit-scrollbar");
    return {
      overflowY: style.overflowY,
      scrollbarWidth: style.scrollbarWidth,
      webkitDisplay: webkitScrollbar.display,
      webkitWidth: webkitScrollbar.width,
    };
  });


  await page.keyboard.press("Escape");


  await page.waitForFunction(
    () =>
      document.querySelector('.sheet-cell[data-address="B14"]')?.getAttribute("contenteditable") === null &&
      Boolean(document.querySelector('.sheet-cell[data-address="B14"] .cell-content')),
    null,
    { timeout: 10000 },
  );


	  const editExitRender = await page.evaluate(() => {
	    const cell = document.querySelector('.sheet-cell[data-address="B14"]');
	    const content = cell?.querySelector(".cell-content");
	    return {
      editable: cell?.getAttribute("contenteditable"),
      hasContentWrapper: Boolean(content),
      directTextNode: [...(cell?.childNodes || [])].some((node) => node.nodeType === Node.TEXT_NODE && node.textContent.trim()),
      contentOpacity: content ? getComputedStyle(content).opacity : "",
	      text: content?.textContent || "",
	    };
	  });


	  const multilineCommitDisplay = await page.evaluate(async () => {
	    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
	    if (!window.__webSheetDevSetCell || !window.__webSheetDevWorkbookState) return null;
	    const row = 30;
	    const col = 20;
	    const address = "T30";
	    const sheetName = window.__webSheetDevWorkbookState().activeSheetName;
	    const beforeSheet = window.__webSheetDevSheetModel(sheetName);
	    const beforeHeight = Number(beforeSheet?.rows?.[row - 1]?.height || 0);
	    window.__webSheetDevSetCell(row, col, "上\n下");
	    await frame();
	    const afterSheet = window.__webSheetDevSheetModel(sheetName);
	    const model = window.__webSheetDevCellModel(sheetName, row, col);
	    const result = {
	      address,
	      raw: model?.raw,
	      display: window.__webSheetDevCellDisplay(sheetName, row, col),
	      whiteSpace: model?.css?.whiteSpace || "",
	      overflowWrap: model?.css?.overflowWrap || "",
	      beforeHeight,
	      afterHeight: Number(afterSheet?.rows?.[row - 1]?.height || 0),
	    };
	    window.__webSheetDevSelectRange?.(1, 1, 1, 1);
	    const host = document.querySelector("#sheetHost");
	    if (host) {
	      host.scrollTop = 0;
	      host.scrollLeft = 0;
	      host.dispatchEvent(new Event("scroll", { bubbles: true }));
	    }
	    await frame();
	    return result;
	  });



	  const formulaEmptyAddress = await page.evaluate(() => {
    const reserved = new Set(["B2", "B14", "D14", "G6", "H6", "I6", "J6"]);
    return [...document.querySelectorAll('.sheet-cell[data-address]')]
      .map((cell) => ({
        address: cell.dataset.address,
        text: cell.querySelector(".cell-content")?.textContent || "",
      }))
      .find((item) => !reserved.has(item.address) && item.text === "")?.address;
  });


  if (!formulaEmptyAddress) throw new Error("No visible empty cell was available for formula bar testing.");


  await page.locator(`.sheet-cell[data-address="${formulaEmptyAddress}"]`).click();


  const formulaBarInitial = await page.evaluate(() => ({
    placeholder: document.querySelector("#formulaInput")?.getAttribute("placeholder"),
    hasApplyButton: Boolean(document.querySelector("#applyFormulaButton")),
    value: document.querySelector("#formulaInput")?.value,
  }));


  const readFormulaBarSizing = () =>
    page.evaluate(() => {
      const bar = document.querySelector(".formula-bar");
      const input = document.querySelector("#formulaInput");
      const button = document.querySelector("#formulaExpandButton");
      const barRect = bar?.getBoundingClientRect();
      const inputRect = input?.getBoundingClientRect();
      const style = input ? getComputedStyle(input) : null;
      return {
        rows: Number(input?.getAttribute("rows") || 0),
        rowHeight: Math.round(barRect?.height || 0),
        inputHeight: Math.round(inputRect?.height || 0),
        fontSize: style?.fontSize,
        lineHeight: Number.parseFloat(style?.lineHeight || "0"),
        expanded: button?.getAttribute("aria-expanded"),
        disabled: button?.disabled,
      };
    });


  const formulaBarSizing = { initial: await readFormulaBarSizing() };


  await page.click("#formulaExpandButton");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.getAttribute("rows") === "3", null, { timeout: 10000 });


  formulaBarSizing.expanded3 = await readFormulaBarSizing();


  await page.click("#formulaExpandButton");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.getAttribute("rows") === "1", null, { timeout: 10000 });


  formulaBarSizing.collapsedAfter3 = await readFormulaBarSizing();


  const formulaHandleBox = await page.locator("#formulaResizeHandle").boundingBox();


  if (!formulaHandleBox) throw new Error("Formula bar resize handle was not visible.");


  const dragLineHeight = formulaBarSizing.initial.lineHeight || 18;


  await page.mouse.move(formulaHandleBox.x + formulaHandleBox.width / 2, formulaHandleBox.y + formulaHandleBox.height / 2);


  await page.mouse.down();


  await page.mouse.move(formulaHandleBox.x + formulaHandleBox.width / 2, formulaHandleBox.y + formulaHandleBox.height / 2 + dragLineHeight * 4, { steps: 8 });


  await page.mouse.up();


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.getAttribute("rows") === "5", null, { timeout: 10000 });


  formulaBarSizing.manual5 = await readFormulaBarSizing();


  await page.click("#formulaExpandButton");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.getAttribute("rows") === "1", null, { timeout: 10000 });


  formulaBarSizing.collapsedAfterManual = await readFormulaBarSizing();


  await page.click("#formulaExpandButton");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.getAttribute("rows") === "5", null, { timeout: 10000 });


  formulaBarSizing.reopenedManual5 = await readFormulaBarSizing();


  await page.click("#formulaExpandButton");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.getAttribute("rows") === "1", null, { timeout: 10000 });


  formulaBarSizing.restored = await readFormulaBarSizing();


  await page.click("#formulaExpandButton");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.getAttribute("rows") === "5", null, { timeout: 10000 });


  await page.locator("#formulaInput").evaluate((input) => {
    input.value = Array.from({ length: 24 }, (_item, index) => `Line ${index + 1}`).join("\n");
    input.scrollTop = 0;
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });


  const formulaScrollInputBox = await page.locator("#formulaInput").boundingBox();


  if (!formulaScrollInputBox) throw new Error("Formula bar input was not visible for scroll testing.");


  await page.mouse.move(formulaScrollInputBox.x + 12, formulaScrollInputBox.y + 8);


  await page.mouse.down();


  await page.mouse.move(formulaScrollInputBox.x + 12, formulaScrollInputBox.y + formulaScrollInputBox.height + 54, { steps: 8 });


  await page.waitForTimeout(260);


  await page.mouse.up();


  formulaBarSizing.hiddenScroll = await page.evaluate(() => {
    const input = document.querySelector("#formulaInput");
    const style = input ? getComputedStyle(input) : null;
    return {
      scrollTop: Math.round(input?.scrollTop || 0),
      scrollHeight: Math.round(input?.scrollHeight || 0),
      clientHeight: Math.round(input?.clientHeight || 0),
      overflowY: style?.overflowY,
      scrollbarWidth: style?.scrollbarWidth,
      webkitDisplay: getComputedStyle(input, "::-webkit-scrollbar")?.display,
      buttonBorderTopWidth: getComputedStyle(document.querySelector("#formulaExpandButton"))?.borderTopWidth,
      buttonBackground: getComputedStyle(document.querySelector("#formulaExpandButton"))?.backgroundColor,
    };
  });


  await page.click("#formulaExpandButton");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.getAttribute("rows") === "1", null, { timeout: 10000 });


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("FormulaBarEdit");


  const formulaBarDuringEdit = await page.evaluate((address) => {
    const cell = document.querySelector(`.sheet-cell[data-address="${address}"]`);
    return {
      contenteditable: cell?.getAttribute("contenteditable"),
      previewText: cell?.querySelector(".cell-content")?.textContent,
      inputValue: document.querySelector("#formulaInput")?.value,
    };
  }, formulaEmptyAddress);


  await page.keyboard.press("Escape");


  await page.waitForFunction((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent === "", formulaEmptyAddress, { timeout: 10000 });


  const formulaBarEscape = await page.evaluate((address) => {
    const cell = document.querySelector(`.sheet-cell[data-address="${address}"]`);
    return {
      contenteditable: cell?.getAttribute("contenteditable"),
      text: cell?.querySelector(".cell-content")?.textContent,
      inputValue: document.querySelector("#formulaInput")?.value,
      status: document.querySelector("#statusText")?.textContent,
    };
  }, formulaEmptyAddress);


  await page.locator("#formulaInput").click();


  await page.keyboard.type("FormulaBarEdit");


  await page.keyboard.press("Enter");


  await page.waitForFunction((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent === "FormulaBarEdit", formulaEmptyAddress, { timeout: 10000 });


  const formulaBarEnter = await page.evaluate((address) => {
    const cell = document.querySelector(`.sheet-cell[data-address="${address}"]`);
    return {
      contenteditable: cell?.getAttribute("contenteditable"),
      text: cell?.querySelector(".cell-content")?.textContent,
      inputValue: document.querySelector("#formulaInput")?.value,
      status: document.querySelector("#statusText")?.textContent,
    };
  }, formulaEmptyAddress);


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=su");


  await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });


  const formulaAutocompleteBarOpen = await page.evaluate(() => ({
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 5).map((item) => item.textContent),
    syntaxes: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-syntax")].slice(0, 5).map((item) => item.textContent),
    count: document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]").length,
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
    role: document.querySelector("#formulaAutocomplete")?.getAttribute("role") || "",
    ariaLabel: document.querySelector("#formulaAutocomplete")?.getAttribute("aria-label") || "",
    activeDescendant: document.querySelector("#formulaAutocomplete")?.getAttribute("aria-activedescendant") || "",
    activeId: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active")?.id || "",
    selectedAria: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active")?.getAttribute("aria-selected") || "",
    inputValue: document.querySelector("#formulaInput")?.value || "",
  }));


  const formulaAutocompleteImeEnter = await page.evaluate(async () => {
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const input = document.querySelector("#formulaInput");
    if (!input) return null;
    const compositionEvent = (type, data) => {
      if (typeof CompositionEvent === "function") return new CompositionEvent(type, { bubbles: true, data });
      const event = new Event(type, { bubbles: true });
      Object.defineProperty(event, "data", { value: data });
      return event;
    };
    input.dispatchEvent(compositionEvent("compositionstart", "su"));
    const enterEvent = new KeyboardEvent("keydown", {
      key: "Enter",
      code: "Enter",
      bubbles: true,
      cancelable: true,
      isComposing: true,
    });
    Object.defineProperty(enterEvent, "keyCode", { get: () => 229 });
    Object.defineProperty(enterEvent, "which", { get: () => 229 });
    input.dispatchEvent(enterEvent);
    await frame();
    const result = {
      defaultPrevented: enterEvent.defaultPrevented,
      inputValue: input.value || "",
      popupVisible: Boolean(document.querySelector("#formulaAutocomplete")),
      active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
    };
    input.dispatchEvent(compositionEvent("compositionend", "su"));
    await frame();
    return result;
  });


  await page.keyboard.press("End");


  const formulaAutocompleteBarKeyboardEnd = await page.evaluate(() => ({
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
    activeDescendant: document.querySelector("#formulaAutocomplete")?.getAttribute("aria-activedescendant") || "",
    activeId: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active")?.id || "",
  }));


  await page.keyboard.press("Home");


  const formulaAutocompleteBarKeyboardHome = await page.evaluate(() => ({
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
    activeDescendant: document.querySelector("#formulaAutocomplete")?.getAttribute("aria-activedescendant") || "",
    activeId: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active")?.id || "",
  }));


  await page.keyboard.press("PageDown");


  const formulaAutocompleteBarKeyboardPageDown = await page.evaluate(() => ({
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
    activeDescendant: document.querySelector("#formulaAutocomplete")?.getAttribute("aria-activedescendant") || "",
    activeId: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active")?.id || "",
  }));


  await page.keyboard.press("PageUp");


  const formulaAutocompleteBarKeyboardPageUp = await page.evaluate(() => ({
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
    activeDescendant: document.querySelector("#formulaAutocomplete")?.getAttribute("aria-activedescendant") || "",
    activeId: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active")?.id || "",
  }));


  await page.keyboard.press("ArrowDown");


  await page.keyboard.press("ArrowUp");


  await page.keyboard.press("Tab");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SUM(", null, { timeout: 5000 });


  await page.waitForSelector("#formulaAutocomplete .formula-autocomplete-hint", { timeout: 5000 });


  const formulaAutocompleteBarAccepted = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    hint: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint")?.textContent?.replace(/\s+/g, " ").trim() || "",
    activeArgument: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint .is-active")?.textContent || "",
    activeDescription: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint-detail")?.textContent || "",
    listVisible: Boolean(document.querySelector("#formulaAutocomplete [data-formula-autocomplete-item]")),
    previewText: document.querySelector(`.sheet-cell[data-address="${document.querySelector("#selectedCellLabel")?.textContent?.trim()}"]`)?.textContent || "",
  }));


  await page.keyboard.type("2,3,");


  await page.waitForFunction(() =>
    document.querySelector("#formulaInput")?.value === "=SUM(2,3," &&
    /number3/.test(document.querySelector("#formulaAutocomplete .formula-autocomplete-hint .is-active")?.textContent || ""),
  null, { timeout: 5000 });


  const formulaAutocompleteBarRepeatHint = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    hint: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint")?.textContent?.replace(/\s+/g, " ").trim() || "",
    activeArgument: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint .is-active")?.textContent || "",
    activeDescription: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint-detail")?.textContent || "",
    listVisible: Boolean(document.querySelector("#formulaAutocomplete [data-formula-autocomplete-item]")),
  }));


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SUM({1,2},");


  await page.waitForFunction(() =>
    document.querySelector("#formulaInput")?.value === "=SUM({1,2}," &&
    /number2/.test(document.querySelector("#formulaAutocomplete .formula-autocomplete-hint .is-active")?.textContent || ""),
  null, { timeout: 5000 });


  const formulaAutocompleteBarArrayHint = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    hint: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint")?.textContent?.replace(/\s+/g, " ").trim() || "",
    activeArgument: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint .is-active")?.textContent || "",
    activeDescription: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint-detail")?.textContent || "",
    listVisible: Boolean(document.querySelector("#formulaAutocomplete [data-formula-autocomplete-item]")),
  }));


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=if");


  await page.keyboard.press("Control+Shift+A");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=IF(logical_test, value_if_true, value_if_false)", null, { timeout: 5000 });


  await page.waitForFunction(() =>
    String(document.querySelector("#formulaAutocomplete .formula-autocomplete-hint-detail")?.textContent || "").includes("logical_test"),
  null, { timeout: 5000 });


  const formulaAutocompleteBarArgumentExpansion = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    hint: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint")?.textContent?.replace(/\s+/g, " ").trim() || "",
    activeDescription: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint-detail")?.textContent || "",
  }));


  await page.locator("#formulaInput").focus();


  await page.keyboard.press("Control+A");


  await page.waitForSelector('[data-data-dialog="function-arguments"]', { timeout: 5000 });


  const formulaAutocompleteBarFunctionArgumentsInitial = await page.evaluate(() => ({
    title: document.querySelector("#functionArgumentsTitle")?.textContent || "",
    functionName: document.querySelector(".function-argument-heading strong")?.textContent || "",
    labels: [...document.querySelectorAll(".function-argument-name")].map((item) => item.textContent.trim().replace(/:$/u, "")),
    values: [...document.querySelectorAll("[data-formula-argument-input]")].map((input) => input.value),
    hasBack: Boolean(document.querySelector('[data-data-dialog-action="function-arguments-back"]')),
  }));


  await page.fill('[data-formula-argument-input][data-formula-argument-index="0"]', "1=1");


  await page.fill('[data-formula-argument-input][data-formula-argument-index="1"]', "9");


  await page.fill('[data-formula-argument-input][data-formula-argument-index="2"]', "3");


  await page.click('[data-data-dialog="function-arguments"] [data-data-dialog-action="function-arguments-ok"]');


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=IF(1=1,9,3)", null, { timeout: 5000 });


  const formulaAutocompleteBarFunctionArgumentsApplied = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    activeElement: document.activeElement?.id || document.activeElement?.className || document.activeElement?.tagName || "",
    dialogClosed: !document.querySelector('[data-data-dialog="function-arguments"]'),
    previewText: document.querySelector(`.sheet-cell[data-address="${document.querySelector("#selectedCellLabel")?.textContent?.trim()}"]`)?.textContent || "",
  }));


  const clearFormulaAutocompleteDom = async () => {
    await page.evaluate(() => {
      document.querySelector("#formulaAutocomplete")?.remove();
    });
  };


  const closeFormulaAutocompletePopup = async () => {
    await page.keyboard.press("Escape");
    await clearFormulaAutocompleteDom();
    await page.waitForFunction(() => !document.querySelector("#formulaAutocomplete"), null, { timeout: 5000 });
  };


  await closeFormulaAutocompletePopup();


  await page.keyboard.press("Escape");


  await page.waitForFunction((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent === "FormulaBarEdit", formulaEmptyAddress, { timeout: 10000 });


  await page.evaluate(() => {
    window.__webSheetDevAddDefinedName?.({ name: "SalesDataAmounts", refersTo: "=CHT!$B$2:$B$4" });
    window.__webSheetDevAddDefinedName?.({ name: "OtherSheetOnly", refersTo: "=AI製品紹介!$A$1", scope: "sheet", sheetIndex: 1, comment: "別シートのローカル名" });
    window.__webSheetDevAddTable?.("CHT", {
      name: "SalesTable",
      range: { top: 80, left: 30, bottom: 83, right: 33 },
      columns: [{ name: "地域" }, { name: "売上" }, { name: "担当" }, { name: "倍額" }],
    });
  });


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SalesD");


  await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });


  const formulaAutocompleteNameOpen = await page.evaluate(() => ({
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 8).map((item) => item.textContent),
    categories: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-category")].slice(0, 8).map((item) => item.textContent),
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
  }));


  await page.keyboard.press("Tab");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesDataAmounts", null, { timeout: 5000 });


  const formulaAutocompleteNameAccepted = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    popupVisible: Boolean(document.querySelector("#formulaAutocomplete")),
  }));


  if (formulaAutocompleteNameAccepted.popupVisible) {
    await closeFormulaAutocompletePopup();
  }


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=LocalM");


  await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });


  const formulaAutocompleteLocalNameScope = await page.evaluate(() => ({
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 8).map((item) => item.textContent),
    categories: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-category")].slice(0, 8).map((item) => item.textContent),
    descriptions: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-description")].slice(0, 8).map((item) => item.textContent),
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
  }));


  await closeFormulaAutocompletePopup();


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=OtherSheet");


  await page.waitForTimeout(220);


  const formulaAutocompleteOtherSheetLocalName = await page.evaluate(() => ({
    popupVisible: Boolean(document.querySelector("#formulaAutocomplete")),
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 8).map((item) => item.textContent),
  }));


  await closeFormulaAutocompletePopup();


  await page.locator("#formulaInput").focus();


  await page.keyboard.press("Enter");


  await clearFormulaAutocompleteDom();


  await page.waitForFunction(() => !document.querySelector("#formulaAutocomplete"), null, { timeout: 5000 });


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SalesT");


  await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });


  const formulaAutocompleteTableOpen = await page.evaluate(() => ({
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 8).map((item) => item.textContent),
    categories: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-category")].slice(0, 8).map((item) => item.textContent),
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
  }));


  await page.keyboard.press("Enter");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable", null, { timeout: 5000 });


  const formulaAutocompleteTableAccepted = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    popupVisible: Boolean(document.querySelector("#formulaAutocomplete")),
  }));


  if (formulaAutocompleteTableAccepted.popupVisible) {
    await closeFormulaAutocompletePopup();
  }


  await page.locator("#formulaInput").focus();


  await page.keyboard.press("Enter");


  await clearFormulaAutocompleteDom();


  await page.waitForFunction(() => !document.querySelector("#formulaAutocomplete"), null, { timeout: 5000 });


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SalesTable[");


  await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });


  const formulaAutocompleteStructuredOpen = await page.evaluate(() => ({
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 12).map((item) => item.textContent),
    categories: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-category")].slice(0, 12).map((item) => item.textContent),
  }));


  await page.evaluate(() => {
    const button = [...document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]")]
      .find((item) => item.querySelector(".formula-autocomplete-name")?.textContent === "売上");
    button?.click();
  });


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[売上]", null, { timeout: 5000 });


  const formulaAutocompleteStructuredAccepted = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    popupVisible: Boolean(document.querySelector("#formulaAutocomplete")),
  }));


  if (formulaAutocompleteStructuredAccepted.popupVisible) {
    await closeFormulaAutocompletePopup();
  }


  await page.locator("#formulaInput").focus();


  await page.keyboard.press("Enter");


  await clearFormulaAutocompleteDom();


  await page.waitForFunction(() => !document.querySelector("#formulaAutocomplete"), null, { timeout: 5000 });


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SalesTable[[#Data],[");


  await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });


  const formulaAutocompleteStructuredNestedOpen = await page.evaluate(() => ({
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 12).map((item) => item.textContent),
    categories: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-category")].slice(0, 12).map((item) => item.textContent),
  }));


  await page.evaluate(() => {
    const button = [...document.querySelectorAll("#formulaAutocomplete [data-formula-autocomplete-item]")]
      .find((item) => item.querySelector(".formula-autocomplete-name")?.textContent === "売上");
    button?.click();
  });


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SalesTable[[#Data],[売上]]", null, { timeout: 5000 });


  const formulaAutocompleteStructuredNestedAccepted = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    popupVisible: Boolean(document.querySelector("#formulaAutocomplete")),
  }));


  if (formulaAutocompleteStructuredNestedAccepted.popupVisible) {
    await closeFormulaAutocompletePopup();
  }


  await page.locator("#formulaInput").focus();


  await page.keyboard.press("Enter");


  await clearFormulaAutocompleteDom();


  await page.waitForFunction(() => !document.querySelector("#formulaAutocomplete"), null, { timeout: 5000 });


  const formulaAutocompleteEngineDefinitions = await page.evaluate(() => {
    const definitions = window.__webSheetDevFormulaAutocompleteDefinitions?.() || [];
    const wanted = [
      "SUMPRODUCT", "COUNTBLANK", "NETWORKDAYS.INTL", "IMABS", "CEILING.MATH", "AGGREGATE", "LET", "MAKEARRAY", "BYROW", "MAP",
      "ROMAN", "POWER", "SQRT", "F.DIST.RT", "F.INV.RT", "LOGNORM.DIST", "LOGNORM.INV", "NEGBINOM.DIST", "DELTA", "GESTEP",
      "ERFC.PRECISE", "TBILLEQ", "ISPMT",
    ];
    return Object.fromEntries(wanted.map((name) => [
      name,
      definitions.find((definition) => definition.name === name) || null,
    ]));
  });


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=sumpr");


  await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });


  const formulaAutocompleteEngineOpen = await page.evaluate(() => ({
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 8).map((item) => item.textContent),
    categories: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-category")].slice(0, 8).map((item) => item.textContent),
    syntaxes: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-syntax")].slice(0, 8).map((item) => item.textContent),
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
  }));


  await page.keyboard.press("Tab");


  await page.waitForFunction(() => document.querySelector("#formulaInput")?.value === "=SUMPRODUCT(", null, { timeout: 5000 });


  await page.waitForSelector("#formulaAutocomplete .formula-autocomplete-hint", { timeout: 5000 });


  const formulaAutocompleteEngineAccepted = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    hint: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint")?.textContent?.replace(/\s+/g, " ").trim() || "",
  }));


  await clearFormulaAutocompleteDom();


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=lognorm.");


  await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });


  const formulaAutocompleteAdvancedOpen = await page.evaluate(() => ({
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 8).map((item) => item.textContent),
    categories: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-category")].slice(0, 8).map((item) => item.textContent),
    syntaxes: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-syntax")].slice(0, 8).map((item) => item.textContent),
    descriptions: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-description")].slice(0, 8).map((item) => item.textContent),
  }));


  await closeFormulaAutocompletePopup();


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=delta");


  await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });


  const formulaAutocompleteEngineeringOpen = await page.evaluate(() => ({
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 8).map((item) => item.textContent),
    categories: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-category")].slice(0, 8).map((item) => item.textContent),
    syntaxes: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-syntax")].slice(0, 8).map((item) => item.textContent),
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
  }));


  await closeFormulaAutocompletePopup();


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.type("FormulaBarEdit");


  await page.keyboard.press("Enter");


  await page.waitForFunction((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent === "FormulaBarEdit", formulaEmptyAddress, { timeout: 10000 });


  const formulaAutocompleteCellAddress = await page.evaluate((reservedAddress) => {
    const reserved = new Set([reservedAddress, "B2", "B14", "D14", "G6", "H6", "I6", "J6"]);
    return [...document.querySelectorAll('.sheet-cell[data-address]')]
      .map((cell) => ({
        address: cell.dataset.address,
        text: cell.querySelector(".cell-content")?.textContent || "",
      }))
      .find((item) => !reserved.has(item.address) && item.text === "")?.address || "";
  }, formulaEmptyAddress);


  if (!formulaAutocompleteCellAddress) throw new Error("No visible empty cell was available for formula autocomplete testing.");


  await page.locator(`.sheet-cell[data-address="${formulaAutocompleteCellAddress}"]`).dblclick();


  await page.keyboard.type("=av");


  await page.waitForSelector("#formulaAutocomplete [data-formula-autocomplete-item]", { timeout: 5000 });


  const formulaAutocompleteCellOpen = await page.evaluate((address) => ({
    address,
    editable: document.querySelector(`.sheet-cell[data-address="${address}"]`)?.getAttribute("contenteditable"),
    text: document.querySelector(`.sheet-cell[data-address="${address}"]`)?.textContent || "",
    names: [...document.querySelectorAll("#formulaAutocomplete .formula-autocomplete-name")].slice(0, 4).map((item) => item.textContent),
    active: document.querySelector("#formulaAutocomplete .formula-autocomplete-item.is-active .formula-autocomplete-name")?.textContent || "",
  }), formulaAutocompleteCellAddress);


  await page.keyboard.press("Enter");


  await page.waitForFunction((address) => document.querySelector(`.sheet-cell[data-address="${address}"]`)?.textContent === "=AVERAGE(", formulaAutocompleteCellAddress, { timeout: 5000 });


  await page.waitForSelector("#formulaAutocomplete .formula-autocomplete-hint", { timeout: 5000 });


  const formulaAutocompleteCellAccepted = await page.evaluate((address) => ({
    editable: document.querySelector(`.sheet-cell[data-address="${address}"]`)?.getAttribute("contenteditable"),
    text: document.querySelector(`.sheet-cell[data-address="${address}"]`)?.textContent || "",
    hint: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint")?.textContent?.replace(/\s+/g, " ").trim() || "",
  }), formulaAutocompleteCellAddress);


  await page.locator(`.sheet-cell[data-address="${formulaAutocompleteCellAddress}"][contenteditable="true"]`).focus();


  await page.keyboard.press("Control+Shift+A");


  await page.waitForFunction((address) =>
    document.querySelector(`.sheet-cell[data-address="${address}"]`)?.textContent === "=AVERAGE(number1, number2, ...)",
  formulaAutocompleteCellAddress, { timeout: 5000 });


  await page.waitForFunction(() =>
    String(document.querySelector("#formulaAutocomplete .formula-autocomplete-hint-detail")?.textContent || "").includes("number1"),
  null, { timeout: 5000 });


  const formulaAutocompleteCellArgumentExpansion = await page.evaluate((address) => ({
    editable: document.querySelector(`.sheet-cell[data-address="${address}"]`)?.getAttribute("contenteditable"),
    text: document.querySelector(`.sheet-cell[data-address="${address}"]`)?.textContent || "",
    hint: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint")?.textContent?.replace(/\s+/g, " ").trim() || "",
    activeDescription: document.querySelector("#formulaAutocomplete .formula-autocomplete-hint-detail")?.textContent || "",
  }), formulaAutocompleteCellAddress);


  await page.evaluate((address) => {
    const cell = document.querySelector(`.sheet-cell[data-address="${address}"][contenteditable="true"]`);
    if (!cell) return;
    cell.textContent = "=AVERAGE(2,4)";
    cell.dispatchEvent(new Event("input", { bubbles: true }));
    const textNode = cell.firstChild;
    const range = document.createRange();
    range.setStart(textNode || cell, textNode ? cell.textContent.length : cell.childNodes.length);
    range.collapse(true);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    cell.focus({ preventScroll: true });
  }, formulaAutocompleteCellAddress);


  await page.keyboard.press("Control+A");


  await page.waitForSelector('[data-data-dialog="function-arguments"]', { timeout: 5000 });


  const formulaAutocompleteCellFunctionArgumentsInitial = await page.evaluate(() => ({
    title: document.querySelector("#functionArgumentsTitle")?.textContent || "",
    functionName: document.querySelector(".function-argument-heading strong")?.textContent || "",
    labels: [...document.querySelectorAll(".function-argument-name")].map((item) => item.textContent.trim().replace(/:$/u, "")),
    values: [...document.querySelectorAll("[data-formula-argument-input]")].map((input) => input.value),
    hasBack: Boolean(document.querySelector('[data-data-dialog-action="function-arguments-back"]')),
  }));


  await page.fill('[data-formula-argument-input][data-formula-argument-index="1"]', "8");


  await page.click('[data-data-dialog="function-arguments"] [data-data-dialog-action="function-arguments-ok"]');


  await page.waitForFunction((address) =>
    document.querySelector(`.sheet-cell[data-address="${address}"]`)?.textContent === "=AVERAGE(2,8)",
  formulaAutocompleteCellAddress, { timeout: 5000 });


  const formulaAutocompleteCellFunctionArgumentsApplied = await page.evaluate((address) => ({
    editable: document.querySelector(`.sheet-cell[data-address="${address}"]`)?.getAttribute("contenteditable"),
    text: document.querySelector(`.sheet-cell[data-address="${address}"]`)?.textContent || "",
    dialogClosed: !document.querySelector('[data-data-dialog="function-arguments"]'),
    activeElement: document.activeElement?.dataset?.address || document.activeElement?.id || document.activeElement?.className || document.activeElement?.tagName || "",
  }), formulaAutocompleteCellAddress);


  await page.keyboard.press("Escape");


  await page.waitForFunction(() => !document.querySelector("#formulaAutocomplete"), null, { timeout: 5000 });


  await page.keyboard.press("Escape");


  await page.waitForFunction((address) => document.querySelector(`.sheet-cell[data-address="${address}"]`)?.getAttribute("contenteditable") === null, formulaAutocompleteCellAddress, { timeout: 10000 });


  const formulaAutocomplete = {
    barOpen: formulaAutocompleteBarOpen,
    imeEnter: formulaAutocompleteImeEnter,
    keyboard: {
      end: formulaAutocompleteBarKeyboardEnd,
      home: formulaAutocompleteBarKeyboardHome,
      pageDown: formulaAutocompleteBarKeyboardPageDown,
      pageUp: formulaAutocompleteBarKeyboardPageUp,
    },
    barAccepted: formulaAutocompleteBarAccepted,
    barRepeatHint: formulaAutocompleteBarRepeatHint,
    barArrayHint: formulaAutocompleteBarArrayHint,
    nameOpen: formulaAutocompleteNameOpen,
    nameAccepted: formulaAutocompleteNameAccepted,
    localNameScope: formulaAutocompleteLocalNameScope,
    otherSheetLocalName: formulaAutocompleteOtherSheetLocalName,
    tableOpen: formulaAutocompleteTableOpen,
    tableAccepted: formulaAutocompleteTableAccepted,
    structuredOpen: formulaAutocompleteStructuredOpen,
    structuredAccepted: formulaAutocompleteStructuredAccepted,
    structuredNestedOpen: formulaAutocompleteStructuredNestedOpen,
    structuredNestedAccepted: formulaAutocompleteStructuredNestedAccepted,
    engineDefinitions: formulaAutocompleteEngineDefinitions,
    engineOpen: formulaAutocompleteEngineOpen,
    engineAccepted: formulaAutocompleteEngineAccepted,
    advancedOpen: formulaAutocompleteAdvancedOpen,
    engineeringOpen: formulaAutocompleteEngineeringOpen,
    cellOpen: formulaAutocompleteCellOpen,
    cellAccepted: formulaAutocompleteCellAccepted,
    argumentExpansion: {
      bar: formulaAutocompleteBarArgumentExpansion,
      cell: formulaAutocompleteCellArgumentExpansion,
    },
    functionArgumentsShortcut: {
      barInitial: formulaAutocompleteBarFunctionArgumentsInitial,
      barApplied: formulaAutocompleteBarFunctionArgumentsApplied,
      cellInitial: formulaAutocompleteCellFunctionArgumentsInitial,
      cellApplied: formulaAutocompleteCellFunctionArgumentsApplied,
    },
    cellAddress: formulaAutocompleteCellAddress,
  };



  const formulaCustomFunctions = await page.evaluate(() => {
    window.__webSheetDevSetActiveSheet?.("CHT");
    const seed = [
      [90, 27, 1],
      [91, 27, 2],
      [92, 27, 2],
      [93, 27, 4],
      [94, 27, 5],
      [90, 28, 10],
      [91, 28, 20],
      [92, 28, 30],
      [93, 28, 40],
      [94, 28, 50],
      [90, 29, "A"],
      [91, 29, "A"],
      [92, 29, "B"],
      [93, 29, "B"],
      [94, 29, "A"],
      [90, 30, "x"],
      [91, 30, ""],
      [92, 30, "y"],
      [93, 30, ""],
      [94, 30, ""],
      [95, 27, 1],
      [95, 28, 0],
      [90, 37, "Region"],
      [90, 38, "Sales"],
      [90, 39, "Rep"],
      [91, 37, "East"],
      [91, 38, 100],
      [91, 39, "A"],
      [92, 37, "West"],
      [92, 38, 200],
      [92, 39, "B"],
      [93, 37, "East"],
      [93, 38, 150],
      [93, 39, "C"],
      [94, 37, "East"],
      [94, 38, "text"],
      [94, 39, "D"],
      [90, 41, "Region"],
      [91, 41, "East"],
    ];
    seed.forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
    window.__webSheetDevAddDefinedName?.({ name: "FormulaSeed", refersTo: "=CHT!$AA$90:$AB$91" });
    const formulas = [
      ["textjoin", 90, 31, '=TEXTJOIN("-",TRUE,AD90:AD94)'],
      ["concat", 91, 31, '=CONCAT("A",AD90:AD94)'],
      ["averageifs", 92, 31, '=AVERAGEIFS(AB90:AB94,AC90:AC94,"A")'],
      ["modeSngl", 93, 31, "=MODE.SNGL(AA90:AA94)"],
      ["rankEq", 94, 31, "=RANK.EQ(4,AA90:AA94,0)"],
      ["percentileInc", 90, 32, "=PERCENTILE.INC(AA90:AA94,0.5)"],
      ["percentileExc", 91, 32, "=PERCENTILE.EXC(AA90:AA94,0.5)"],
      ["quartileInc", 92, 32, "=QUARTILE.INC(AA90:AA94,3)"],
      ["quartileExc", 93, 32, "=QUARTILE.EXC(AA90:AA94,2)"],
      ["intercept", 94, 32, "=INTERCEPT(AB90:AB94,AA90:AA94)"],
      ["forecastLinear", 90, 33, "=FORECAST.LINEAR(3,AB90:AB94,AA90:AA94)"],
      ["dollar", 91, 33, "=DOLLAR(1234.567,2)"],
      ["fixed", 92, 33, "=FIXED(1234.567,1,FALSE)"],
      ["numbervalue", 93, 33, '=NUMBERVALUE("1.234,5",",",".")'],
      ["errorType", 94, 33, "=ERROR.TYPE(1/0)"],
      ["typeText", 90, 34, '=TYPE("abc")'],
      ["typeNumber", 91, 34, "=TYPE(42)"],
      ["typeBoolean", 92, 34, "=TYPE(TRUE)"],
      ["lookup", 90, 35, "=LOOKUP(4,AA90:AA94,AB90:AB94)"],
      ["xmatch", 91, 35, "=XMATCH(4,AA90:AA94)"],
      ["aggregateSum", 92, 36, "=AGGREGATE(9,0,AA90:AA94)"],
      ["aggregateLarge", 93, 36, "=AGGREGATE(14,0,AA90:AA94,2)"],
      ["aggregateMedian", 94, 36, "=AGGREGATE(12,0,AA90:AA94)"],
      ["aggregateIgnoreError", 95, 36, "=AGGREGATE(9,6,AA90:AA94,1/0)"],
      ["ifs", 96, 36, "=IFS(FALSE,1,TRUE,2)"],
      ["xor", 97, 36, "=XOR(TRUE,FALSE,FALSE)"],
      ["letSimple", 98, 36, "=LET(x,2,x+3)"],
      ["letMulti", 99, 36, "=LET(x,2,y,x+3,y*2)"],
      ["letRange", 100, 36, "=LET(nums,AA90:AA94,SUM(nums))"],
      ["makearray", 90, 125, "=MAKEARRAY(2,2,LAMBDA(r,c,r+c))"],
      ["byrow", 90, 128, "=BYROW(AA90:AB91,LAMBDA(row,SUM(row)))"],
      ["bycol", 90, 130, "=BYCOL(AA90:AB91,LAMBDA(col,SUM(col)))"],
      ["map", 90, 133, "=MAP(AA90:AA91,LAMBDA(x,x*2))"],
      ["reduce", 90, 135, "=REDUCE(0,AA90:AA94,LAMBDA(acc,value,acc+value))"],
      ["scan", 90, 137, "=SCAN(0,AA90:AA92,LAMBDA(acc,value,acc+value))"],
      ["sortby", 90, 139, "=SORTBY(AA90:AB94,AB90:AB94,-1)"],
      ["textsplitMultiDelimiter", 90, 142, '=TEXTSPLIT("A,B;C|D",{",",";","|"})'],
      ["xmatchWildcard", 90, 146, '=XMATCH("A*",AC90:AC94,2)'],
      ["xmatchWildcardReverse", 91, 146, '=XMATCH("A*",AC90:AC94,2,-1)'],
      ["takeLastRows", 90, 148, "=TAKE(AA90:AB94,-2)"],
      ["dropLastRows", 90, 151, "=DROP(AA90:AB94,-2)"],
      ["convertLength", 92, 35, '=CONVERT(1,"m","ft")'],
      ["convertTemp", 93, 35, '=CONVERT(32,"F","C")'],
      ["dsum", 90, 43, "=DSUM(AK90:AM94,AL90,AO90:AO91)"],
      ["dcount", 91, 43, "=DCOUNT(AK90:AM94,AL90,AO90:AO91)"],
      ["dcounta", 92, 43, "=DCOUNTA(AK90:AM94,AL90,AO90:AO91)"],
      ["daverage", 93, 43, "=DAVERAGE(AK90:AM94,AL90,AO90:AO91)"],
      ["dmax", 94, 43, "=DMAX(AK90:AM94,AL90,AO90:AO91)"],
      ["dmin", 95, 43, "=DMIN(AK90:AM94,AL90,AO90:AO91)"],
      ["unique", 90, 45, "=UNIQUE(AC90:AC94)"],
      ["sort", 90, 47, "=SORT(AA90:AB94,2,-1,FALSE)"],
      ["filter", 90, 50, "=FILTER(AA90:AB94,AA90:AA94>2)"],
      ["sequence", 90, 53, "=SEQUENCE(2,2,10,5)"],
      ["indirect", 93, 53, '=INDIRECT("AB91")'],
      ["offsetStatic", 94, 53, "=OFFSET(AA90,3,1)"],
      ["offsetRange", 90, 56, "=OFFSET(AA90,1,0,2,2)"],
      ["dproduct", 96, 43, "=DPRODUCT(AK90:AM94,AL90,AO90:AO91)"],
      ["dstdev", 97, 43, "=DSTDEV(AK90:AM94,AL90,AO90:AO91)"],
      ["dstdevp", 98, 43, "=DSTDEVP(AK90:AM94,AL90,AO90:AO91)"],
      ["dvar", 99, 43, "=DVAR(AK90:AM94,AL90,AO90:AO91)"],
      ["dvarp", 100, 43, "=DVARP(AK90:AM94,AL90,AO90:AO91)"],
      ["take", 90, 60, "=TAKE(AA90:AB94,2)"],
      ["drop", 90, 63, "=DROP(AA90:AB94,3)"],
      ["vstack", 90, 66, "=VSTACK(AA90:AB91,AA93:AB94)"],
      ["hstack", 90, 69, "=HSTACK(AA90:AA91,AB90:AB91)"],
      ["chooseRows", 90, 72, "=CHOOSEROWS(AA90:AB94,1,5)"],
      ["chooseCols", 90, 75, "=CHOOSECOLS(AA90:AB94,2)"],
      ["tocol", 90, 77, "=TOCOL(AA90:AB91)"],
      ["torow", 90, 79, "=TOROW(AA90:AB91)"],
      ["wraprows", 90, 83, '=WRAPROWS(AA90:AA94,2,"")'],
      ["wrapcols", 90, 86, '=WRAPCOLS(AA90:AA94,2,"")'],
      ["expand", 90, 90, '=EXPAND(AA90:AB91,3,3,"x")'],
      ["indexColumn", 90, 94, "=INDEX(AA90:AB94,0,2)"],
      ["directRange", 90, 97, "=AA90:AB91"],
      ["indirectR1C1", 90, 100, '=INDIRECT("R90C27:R91C28",FALSE)'],
      ["namedRange", 90, 103, "=FormulaSeed"],
      ["offsetCellArgs", 90, 106, "=OFFSET(AA90,AA95,AB95,2,2)"],
      ["transpose", 90, 109, "=TRANSPOSE(AA90:AB91)"],
      ["chooseRange", 90, 112, "=CHOOSE(2,AA90:AA91,AB90:AB91)"],
      ["textsplit", 90, 115, '=TEXTSPLIT("A,B;C,D",",",";")'],
      ["textbefore", 90, 119, '=TEXTBEFORE("A-B-C","-",2)'],
      ["textafter", 90, 120, '=TEXTAFTER("A-B-C","-",2)'],
      ["textsplitIgnoreEmpty", 90, 121, '=TEXTSPLIT("A,,B",",",,TRUE)'],
      ["textafterLast", 90, 123, '=TEXTAFTER("A-B-C","-",-1)'],
      ["textbeforeFallback", 90, 124, '=TEXTBEFORE("ABC","-",1,0,FALSE,"NF")'],
    ];
    [
      [90, 115], [90, 116], [91, 115], [91, 116],
      [90, 121], [90, 122],
      [90, 125], [90, 126], [91, 125], [91, 126],
      [90, 128], [91, 128],
      [90, 130], [90, 131],
      [90, 133], [91, 133],
      [90, 137], [91, 137], [92, 137],
      [90, 139], [90, 140], [91, 139], [91, 140],
      [90, 142], [90, 143], [90, 144], [90, 145],
      [90, 148], [90, 149], [91, 148], [91, 149],
      [90, 151], [90, 152], [91, 151], [91, 152], [92, 151], [92, 152],
    ].forEach(([row, col]) => window.__webSheetDevSetCell?.(row, col, ""));
    formulas.forEach(([, row, col, formula]) => window.__webSheetDevSetCell?.(row, col, formula));
    const values = Object.fromEntries(formulas.map(([key, row, col]) => {
      const engine = window.__webSheetDevFormulaEngineCell?.("CHT", row, col) || {};
      return [key, {
        serialized: engine.serialized,
        formula: engine.formula,
        value: engine.value,
        display: window.__webSheetDevCellDisplay?.("CHT", row, col),
      }];
    }));
    const spillDisplays = {
      unique: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 45),
        window.__webSheetDevCellDisplay?.("CHT", 91, 45),
      ],
      sort: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 47),
        window.__webSheetDevCellDisplay?.("CHT", 90, 48),
      ],
      filter: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 50),
        window.__webSheetDevCellDisplay?.("CHT", 90, 51),
        window.__webSheetDevCellDisplay?.("CHT", 91, 50),
      ],
      sequence: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 53),
        window.__webSheetDevCellDisplay?.("CHT", 90, 54),
        window.__webSheetDevCellDisplay?.("CHT", 91, 53),
        window.__webSheetDevCellDisplay?.("CHT", 91, 54),
      ],
      offsetRange: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 56),
        window.__webSheetDevCellDisplay?.("CHT", 90, 57),
        window.__webSheetDevCellDisplay?.("CHT", 91, 56),
      ],
      take: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 60),
        window.__webSheetDevCellDisplay?.("CHT", 90, 61),
        window.__webSheetDevCellDisplay?.("CHT", 91, 60),
      ],
      drop: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 63),
        window.__webSheetDevCellDisplay?.("CHT", 90, 64),
        window.__webSheetDevCellDisplay?.("CHT", 91, 63),
      ],
      vstack: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 66),
        window.__webSheetDevCellDisplay?.("CHT", 93, 66),
        window.__webSheetDevCellDisplay?.("CHT", 93, 67),
      ],
      hstack: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 69),
        window.__webSheetDevCellDisplay?.("CHT", 90, 70),
        window.__webSheetDevCellDisplay?.("CHT", 91, 69),
      ],
      chooseRows: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 72),
        window.__webSheetDevCellDisplay?.("CHT", 91, 72),
        window.__webSheetDevCellDisplay?.("CHT", 91, 73),
      ],
      chooseCols: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 75),
        window.__webSheetDevCellDisplay?.("CHT", 91, 75),
      ],
      tocol: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 77),
        window.__webSheetDevCellDisplay?.("CHT", 91, 77),
        window.__webSheetDevCellDisplay?.("CHT", 92, 77),
        window.__webSheetDevCellDisplay?.("CHT", 93, 77),
      ],
      torow: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 79),
        window.__webSheetDevCellDisplay?.("CHT", 90, 80),
        window.__webSheetDevCellDisplay?.("CHT", 90, 81),
        window.__webSheetDevCellDisplay?.("CHT", 90, 82),
      ],
      wraprows: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 83),
        window.__webSheetDevCellDisplay?.("CHT", 90, 84),
        window.__webSheetDevCellDisplay?.("CHT", 92, 83),
        window.__webSheetDevCellDisplay?.("CHT", 92, 84),
      ],
      wrapcols: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 86),
        window.__webSheetDevCellDisplay?.("CHT", 91, 86),
        window.__webSheetDevCellDisplay?.("CHT", 90, 88),
        window.__webSheetDevCellDisplay?.("CHT", 91, 88),
      ],
      expand: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 90),
        window.__webSheetDevCellDisplay?.("CHT", 90, 92),
        window.__webSheetDevCellDisplay?.("CHT", 92, 90),
        window.__webSheetDevCellDisplay?.("CHT", 92, 92),
      ],
      indexColumn: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 94),
        window.__webSheetDevCellDisplay?.("CHT", 91, 94),
      ],
      directRange: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 97),
        window.__webSheetDevCellDisplay?.("CHT", 90, 98),
        window.__webSheetDevCellDisplay?.("CHT", 91, 97),
      ],
      indirectR1C1: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 100),
        window.__webSheetDevCellDisplay?.("CHT", 90, 101),
        window.__webSheetDevCellDisplay?.("CHT", 91, 100),
      ],
      namedRange: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 103),
        window.__webSheetDevCellDisplay?.("CHT", 90, 104),
        window.__webSheetDevCellDisplay?.("CHT", 91, 103),
      ],
      offsetCellArgs: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 106),
        window.__webSheetDevCellDisplay?.("CHT", 90, 107),
        window.__webSheetDevCellDisplay?.("CHT", 91, 106),
      ],
      transpose: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 109),
        window.__webSheetDevCellDisplay?.("CHT", 90, 110),
        window.__webSheetDevCellDisplay?.("CHT", 91, 109),
      ],
      chooseRange: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 112),
        window.__webSheetDevCellDisplay?.("CHT", 91, 112),
      ],
      textsplit: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 115),
        window.__webSheetDevCellDisplay?.("CHT", 90, 116),
        window.__webSheetDevCellDisplay?.("CHT", 91, 115),
        window.__webSheetDevCellDisplay?.("CHT", 91, 116),
      ],
      textsplitIgnoreEmpty: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 121),
        window.__webSheetDevCellDisplay?.("CHT", 90, 122),
      ],
      makearray: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 125),
        window.__webSheetDevCellDisplay?.("CHT", 90, 126),
        window.__webSheetDevCellDisplay?.("CHT", 91, 125),
        window.__webSheetDevCellDisplay?.("CHT", 91, 126),
      ],
      byrow: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 128),
        window.__webSheetDevCellDisplay?.("CHT", 91, 128),
      ],
      bycol: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 130),
        window.__webSheetDevCellDisplay?.("CHT", 90, 131),
      ],
      map: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 133),
        window.__webSheetDevCellDisplay?.("CHT", 91, 133),
      ],
      scan: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 137),
        window.__webSheetDevCellDisplay?.("CHT", 91, 137),
        window.__webSheetDevCellDisplay?.("CHT", 92, 137),
      ],
      sortby: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 139),
        window.__webSheetDevCellDisplay?.("CHT", 90, 140),
        window.__webSheetDevCellDisplay?.("CHT", 91, 139),
        window.__webSheetDevCellDisplay?.("CHT", 91, 140),
      ],
      textsplitMultiDelimiter: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 142),
        window.__webSheetDevCellDisplay?.("CHT", 90, 143),
        window.__webSheetDevCellDisplay?.("CHT", 90, 144),
        window.__webSheetDevCellDisplay?.("CHT", 90, 145),
      ],
      takeLastRows: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 148),
        window.__webSheetDevCellDisplay?.("CHT", 90, 149),
        window.__webSheetDevCellDisplay?.("CHT", 91, 148),
        window.__webSheetDevCellDisplay?.("CHT", 91, 149),
      ],
      dropLastRows: [
        window.__webSheetDevCellDisplay?.("CHT", 90, 151),
        window.__webSheetDevCellDisplay?.("CHT", 90, 152),
        window.__webSheetDevCellDisplay?.("CHT", 91, 151),
        window.__webSheetDevCellDisplay?.("CHT", 91, 152),
        window.__webSheetDevCellDisplay?.("CHT", 92, 151),
        window.__webSheetDevCellDisplay?.("CHT", 92, 152),
      ],
    };
    const definitions = window.__webSheetDevFormulaAutocompleteDefinitions?.() || [];
    const wanted = [
      "TEXTJOIN",
      "CONCAT",
      "TEXTSPLIT",
      "TEXTBEFORE",
      "TEXTAFTER",
      "AVERAGEIFS",
      "MODE.SNGL",
      "RANK.EQ",
      "PERCENTILE.INC",
      "QUARTILE.EXC",
      "FORECAST.LINEAR",
      "INTERCEPT",
      "DOLLAR",
      "FIXED",
      "NUMBERVALUE",
      "ERROR.TYPE",
      "TYPE",
      "LOOKUP",
      "XMATCH",
      "AGGREGATE",
      "IFS",
      "XOR",
      "LET",
      "LAMBDA",
      "MAKEARRAY",
      "BYROW",
      "BYCOL",
      "MAP",
      "REDUCE",
      "SCAN",
      "SORT",
      "SORTBY",
      "UNIQUE",
      "FILTER",
      "SEQUENCE",
      "INDIRECT",
      "TAKE",
      "DROP",
      "VSTACK",
      "HSTACK",
      "CHOOSEROWS",
      "CHOOSECOLS",
      "TOCOL",
      "TOROW",
      "WRAPROWS",
      "WRAPCOLS",
      "EXPAND",
      "CONVERT",
      "DSUM",
      "DCOUNT",
      "DCOUNTA",
      "DAVERAGE",
      "DMAX",
      "DMIN",
      "DPRODUCT",
      "DSTDEV",
      "DSTDEVP",
      "DVAR",
      "DVARP",
    ];
    const engineDefinitions = Object.fromEntries(wanted.map((name) => [
      name,
      definitions.find((definition) => definition.name === name) || null,
    ]));
    return { values, spillDisplays, engineDefinitions };
  });



  const selectFormulaReferenceAddress = async (address) => {
    await page.evaluate((targetAddress) => {
      const match = String(targetAddress || "").match(/^([A-Z]+)(\d+)$/);
      if (!match) return false;
      const col = Array.from(match[1]).reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
      const row = Number(match[2]);
      window.__webSheetDevSelectRange?.(row, col, row, col);
      return true;
    }, address);
    await page.waitForSelector(`.sheet-cell[data-address="${address}"]`, { timeout: 10000 });
  };


  await selectFormulaReferenceAddress(formulaEmptyAddress);



  const formulaReferenceTargets = await page.evaluate(({ editAddress, cellEditAddress }) => {
    const colName = (col) => {
      let value = "";
      let n = Number(col) || 0;
      while (n > 0) {
        const mod = (n - 1) % 26;
        value = String.fromCharCode(65 + mod) + value;
        n = Math.floor((n - mod) / 26);
      }
      return value;
    };
    const expandForMerges = (range) => {
      if (!range) return range;
      const sheetName = document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "";
      const merges = window.__webSheetDevSheetModel?.(sheetName)?.merges || [];
      const expanded = { ...range };
      let changed = true;
      while (changed) {
        changed = false;
        merges.forEach((merge) => {
          if (expanded.bottom < merge.top || expanded.top > merge.bottom || expanded.right < merge.left || expanded.left > merge.right) return;
          const next = {
            top: Math.min(expanded.top, merge.top),
            left: Math.min(expanded.left, merge.left),
            bottom: Math.max(expanded.bottom, merge.bottom),
            right: Math.max(expanded.right, merge.right),
          };
          if (next.top !== expanded.top || next.left !== expanded.left || next.bottom !== expanded.bottom || next.right !== expanded.right) {
            Object.assign(expanded, next);
            changed = true;
          }
        });
      }
      return expanded;
    };
    const rangeLabelFromBounds = (range) => {
      if (!range) return "";
      const start = `${colName(range.left)}${range.top}`;
      const end = `${colName(range.right)}${range.bottom}`;
      return start === end ? start : `${start}:${end}`;
    };
    const hostRect = document.querySelector("#sheetHost")?.getBoundingClientRect();
    const visible = [...document.querySelectorAll(".sheet-cell[data-address][data-row][data-col]")]
      .map((cell) => ({
        address: cell.dataset.address,
        row: Number(cell.dataset.row),
        col: Number(cell.dataset.col),
        rect: cell.getBoundingClientRect(),
      }))
      .filter((item) => {
        if (!item.rect.width || !item.rect.height || item.address === editAddress || item.address === cellEditAddress) return false;
        if (!hostRect) return true;
        return (
          item.rect.left >= hostRect.left &&
          item.rect.right <= hostRect.right &&
          item.rect.top >= hostRect.top &&
          item.rect.bottom <= hostRect.bottom &&
          item.rect.left >= 0 &&
          item.rect.top >= 0 &&
          item.rect.right <= window.innerWidth &&
          item.rect.bottom <= window.innerHeight
        );
      });
    const byAddress = new Map(visible.map((item) => [item.address, item]));
    const click = byAddress.get("B2") || visible.find((item) => item.row >= 2 && item.col >= 2) || visible[0];
    const dragStart = byAddress.get("C2") || visible.find((item) => item.row >= 2 && item.col >= 3) || visible[0];
    const dragEnd = byAddress.get("D3") || visible.find((item) => item.row > dragStart?.row && item.col > dragStart?.col) || visible.find((item) => item.address !== dragStart?.address) || dragStart;
    const range = dragStart && dragEnd
      ? expandForMerges({
        top: Math.min(dragStart.row, dragEnd.row),
        left: Math.min(dragStart.col, dragEnd.col),
        bottom: Math.max(dragStart.row, dragEnd.row),
        right: Math.max(dragStart.col, dragEnd.col),
      })
      : null;
    const rangeLabel = rangeLabelFromBounds(range);
    return {
      click: click?.address || "",
      dragStart: dragStart?.address || "",
      dragEnd: dragEnd?.address || "",
      dragExpected: rangeLabel,
      clickCol: click?.col || 0,
      clickRow: click?.row || 0,
      dragEndCol: dragEnd?.col || 0,
      dragEndRow: dragEnd?.row || 0,
      clickColumnName: click ? colName(click.col) : "",
      dragEndColumnName: dragEnd ? colName(dragEnd.col) : "",
    };
  }, { editAddress: formulaEmptyAddress, cellEditAddress: formulaAutocompleteCellAddress });


  if (!formulaReferenceTargets.click || !formulaReferenceTargets.dragStart || !formulaReferenceTargets.dragEnd) {
    throw new Error("No visible cells were available for formula reference insertion testing.");
  }


  await page.locator(`.sheet-cell[data-address="${formulaEmptyAddress}"]`).click();


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SUM(");


  await page.locator(`.sheet-cell[data-address="${formulaReferenceTargets.click}"]`).click();


  await page.waitForFunction((expected) => document.querySelector("#formulaInput")?.value === `=SUM(${expected}`, formulaReferenceTargets.click, { timeout: 5000 });


  const formulaReferenceBarClick = await page.evaluate((expected) => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    selectedLabel: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
    frameVisible: Boolean(document.querySelector(".sheet-formula-reference-frame")),
    highlighted: document.querySelector(`.sheet-cell[data-address="${expected}"]`)?.classList.contains("is-formula-reference"),
  }), formulaReferenceTargets.click);



  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SUM(");


  const dragStartBox = await page.locator(`.sheet-cell[data-address="${formulaReferenceTargets.dragStart}"]`).boundingBox();


  const dragEndBox = await page.locator(`.sheet-cell[data-address="${formulaReferenceTargets.dragEnd}"]`).boundingBox();


  if (!dragStartBox || !dragEndBox) throw new Error("Formula reference drag cells were not visible.");


  await page.mouse.move(dragStartBox.x + dragStartBox.width / 2, dragStartBox.y + dragStartBox.height / 2);


  await page.mouse.down();


  await page.mouse.move(dragEndBox.x + dragEndBox.width / 2, dragEndBox.y + dragEndBox.height / 2, { steps: 6 });


  await page.mouse.up();


  await page.waitForFunction((expected) => document.querySelector("#formulaInput")?.value === `=SUM(${expected}`, formulaReferenceTargets.dragExpected, { timeout: 5000 });


  const formulaReferenceBarDrag = await page.evaluate((expected) => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    frameVisible: Boolean(document.querySelector(".sheet-formula-reference-frame")),
    highlightedCount: [...document.querySelectorAll(".sheet-cell.is-formula-reference")].length,
    expected,
  }), formulaReferenceTargets.dragExpected);



  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type(`=SUM(${formulaReferenceTargets.click})`);


  await page.evaluate((address) => {
    const input = document.querySelector("#formulaInput");
    const start = `=SUM(${address}`.length - Math.max(1, Math.min(2, String(address || "").length));
    input.focus();
    input.setSelectionRange(start, start);
  }, formulaReferenceTargets.click);


  await page.locator(`.sheet-cell[data-address="${formulaReferenceTargets.dragEnd}"]`).click();


  await page.waitForFunction((expected) => document.querySelector("#formulaInput")?.value === `=SUM(${expected})`, formulaReferenceTargets.dragEnd, { timeout: 5000 });


  const formulaReferenceReplace = await page.evaluate((expected) => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    frameVisible: Boolean(document.querySelector(".sheet-formula-reference-frame")),
    highlighted: document.querySelector(`.sheet-cell[data-address="${expected}"]`)?.classList.contains("is-formula-reference"),
  }), formulaReferenceTargets.dragEnd);



  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SUM(");


  await page.locator(`.column-header[data-col="${formulaReferenceTargets.clickCol}"]`).click();


  await page.waitForFunction((expected) => document.querySelector("#formulaInput")?.value === `=SUM(${expected}`, `${formulaReferenceTargets.clickColumnName}:${formulaReferenceTargets.clickColumnName}`, { timeout: 5000 });


  const formulaReferenceColumnClick = await page.evaluate((expectedClassCol) => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    frameVisible: Boolean(document.querySelector(".sheet-formula-reference-frame")),
    highlightedCount: [...document.querySelectorAll(`.sheet-cell[data-col="${expectedClassCol}"].is-formula-reference`)].length,
  }), formulaReferenceTargets.clickCol);



  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SUM(");


  const headerStartBox = await page.locator(`.column-header[data-col="${formulaReferenceTargets.clickCol}"]`).boundingBox();


  const headerEndBox = await page.locator(`.column-header[data-col="${formulaReferenceTargets.dragEndCol}"]`).boundingBox();


  if (!headerStartBox || !headerEndBox) throw new Error("Formula column reference headers were not visible.");


  await page.mouse.move(headerStartBox.x + headerStartBox.width / 2, headerStartBox.y + headerStartBox.height / 2);


  await page.mouse.down();


  await page.mouse.move(headerEndBox.x + headerEndBox.width / 2, headerEndBox.y + headerEndBox.height / 2, { steps: 4 });


  await page.mouse.up();


  const columnDragExpected = `${formulaReferenceTargets.clickColumnName}:${formulaReferenceTargets.dragEndColumnName}`;


  await page.waitForFunction((expected) => document.querySelector("#formulaInput")?.value === `=SUM(${expected}`, columnDragExpected, { timeout: 5000 });


  const formulaReferenceColumnDrag = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    frameVisible: Boolean(document.querySelector(".sheet-formula-reference-frame")),
    highlightedCount: [...document.querySelectorAll(".sheet-cell.is-formula-reference")].length,
  }));



  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SUM(");


  await page.locator(`.row-header[data-row="${formulaReferenceTargets.clickRow}"]`).click();


  await page.waitForFunction((expected) => document.querySelector("#formulaInput")?.value === `=SUM(${expected}`, `${formulaReferenceTargets.clickRow}:${formulaReferenceTargets.clickRow}`, { timeout: 5000 });


  const formulaReferenceRowClick = await page.evaluate((expectedRow) => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    frameVisible: Boolean(document.querySelector(".sheet-formula-reference-frame")),
    highlightedCount: [...document.querySelectorAll(`.sheet-cell[data-row="${expectedRow}"].is-formula-reference`)].length,
  }), formulaReferenceTargets.clickRow);



  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type("=SUM(");


  const rowStartBox = await page.locator(`.row-header[data-row="${formulaReferenceTargets.clickRow}"]`).boundingBox();


  const rowEndBox = await page.locator(`.row-header[data-row="${formulaReferenceTargets.dragEndRow}"]`).boundingBox();


  if (!rowStartBox || !rowEndBox) throw new Error("Formula row reference headers were not visible.");


  await page.mouse.move(rowStartBox.x + rowStartBox.width / 2, rowStartBox.y + rowStartBox.height / 2);


  await page.mouse.down();


  await page.mouse.move(rowEndBox.x + rowEndBox.width / 2, rowEndBox.y + rowEndBox.height / 2, { steps: 4 });


  await page.mouse.up();


  const rowDragExpected = `${formulaReferenceTargets.clickRow}:${formulaReferenceTargets.dragEndRow}`;


  await page.waitForFunction((expected) => document.querySelector("#formulaInput")?.value === `=SUM(${expected}`, rowDragExpected, { timeout: 5000 });


  const formulaReferenceRowDrag = await page.evaluate(() => ({
    inputValue: document.querySelector("#formulaInput")?.value || "",
    frameVisible: Boolean(document.querySelector(".sheet-formula-reference-frame")),
    highlightedCount: [...document.querySelectorAll(".sheet-cell.is-formula-reference")].length,
  }));



  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type(`=${columnDragExpected}`);


  await page.keyboard.press("F4");


  const formulaReferenceColumnF4 = await page.locator("#formulaInput").inputValue();


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type(`=${rowDragExpected}`);


  await page.keyboard.press("F4");


  const formulaReferenceRowF4 = await page.locator("#formulaInput").inputValue();



  await page.locator("#formulaInput").focus();


  await page.keyboard.press("Escape");


  await page.keyboard.press("Escape");


  await page.evaluate((targetAddress) => {
    const match = String(targetAddress || "").match(/^([A-Z]+)(\d+)$/);
    if (!match) return;
    const col = Array.from(match[1]).reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
    const row = Number(match[2]);
    window.__webSheetDevSetCell?.(row, col, "FormulaBarEdit");
    window.__webSheetDevSelectRange?.(row, col, row, col);
    document.querySelector("#formulaAutocomplete")?.remove();
  }, formulaEmptyAddress);



  await selectFormulaReferenceAddress(formulaAutocompleteCellAddress);


  await page.locator(`.sheet-cell[data-address="${formulaAutocompleteCellAddress}"]`).dblclick();


  await page.keyboard.type("=SUM(");


  await page.locator(`.sheet-cell[data-address="${formulaReferenceTargets.click}"]`).click();


  await page.waitForFunction(
    ({ address, expected }) => document.querySelector(`.sheet-cell[data-address="${address}"]`)?.textContent === `=SUM(${expected}`,
    { address: formulaAutocompleteCellAddress, expected: formulaReferenceTargets.click },
    { timeout: 5000 },
  );


  const formulaReferenceCellClick = await page.evaluate(({ address, expected }) => {
    const cell = document.querySelector(`.sheet-cell[data-address="${address}"]`);
    return {
      editable: cell?.getAttribute("contenteditable"),
      text: cell?.textContent || "",
      activeEditor: document.activeElement === cell,
      frameVisible: Boolean(document.querySelector(".sheet-formula-reference-frame")),
      highlighted: document.querySelector(`.sheet-cell[data-address="${expected}"]`)?.classList.contains("is-formula-reference"),
    };
  }, { address: formulaAutocompleteCellAddress, expected: formulaReferenceTargets.click });


  await page.keyboard.press("Escape");


  await page.keyboard.press("Escape");


  await page.waitForFunction((address) => document.querySelector(`.sheet-cell[data-address="${address}"]`)?.getAttribute("contenteditable") === null, formulaAutocompleteCellAddress, { timeout: 10000 });



  await selectFormulaReferenceAddress(formulaEmptyAddress);


  await page.locator(`.sheet-cell[data-address="${formulaEmptyAddress}"]`).click();


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type(`=${formulaReferenceTargets.click}+${formulaReferenceTargets.dragEnd}`);


  await page.waitForFunction(() => document.querySelectorAll(".sheet-formula-reference-frame").length >= 2, null, { timeout: 5000 });


  const formulaReferenceColors = await page.evaluate(({ first, second }) => {
    const frames = [...document.querySelectorAll(".sheet-formula-reference-frame")];
    const cellClass = (address) => [...(document.querySelector(`.sheet-cell[data-address="${address}"]`)?.classList || [])].find((name) => name.startsWith("formula-reference-color-")) || "";
    return {
      frameCount: frames.length,
      frameClasses: frames.map((frame) => [...frame.classList].find((name) => name.startsWith("formula-reference-color-")) || ""),
      firstCellClass: cellClass(first),
      secondCellClass: cellClass(second),
    };
  }, { first: formulaReferenceTargets.click, second: formulaReferenceTargets.dragEnd });



  const f4Match = formulaReferenceTargets.click.match(/^([A-Z]+)(\d+)$/);


  const formulaReferenceF4Expected = f4Match
    ? [
        `=$${f4Match[1]}$${f4Match[2]}`,
        `=${f4Match[1]}$${f4Match[2]}`,
        `=$${f4Match[1]}${f4Match[2]}`,
        `=${f4Match[1]}${f4Match[2]}`,
      ]
    : [];


  await page.locator("#formulaInput").click();


  await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");


  await page.keyboard.press("Backspace");


  await page.keyboard.type(`=${formulaReferenceTargets.click}`);


  const formulaReferenceF4 = [];


  for (let index = 0; index < 4; index += 1) {
    await page.keyboard.press("F4");
    formulaReferenceF4.push(await page.locator("#formulaInput").inputValue());
  }



  const formulaReferenceCrossSheet = await (async () => {
    const tabs = await page.evaluate(() => [...document.querySelectorAll(".sheet-tab")].map((tab, index) => ({
      index,
      sheetIndex: Number(tab.dataset.sheetIndex),
      name: tab.textContent.trim(),
      active: tab.classList.contains("is-active"),
    })));
    const source = tabs.find((tab) => tab.active);
    const target = tabs.find((tab) => !tab.active);
    if (!source || !target) return { skipped: true };
    const crossSheetSourceAddress = formulaReferenceTargets.dragEnd || formulaAutocompleteCellAddress || formulaEmptyAddress;
    const sourceMatch = crossSheetSourceAddress.match(/^([A-Z]+)(\d+)$/);
    if (!sourceMatch) return { skipped: true };
    const sourceCol = Array.from(sourceMatch[1]).reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
    const sourceRow = Number(sourceMatch[2]);
    await page.evaluate(({ row, col }) => {
      window.__webSheetDevSetCell?.(row, col, "");
      window.__webSheetDevSelectRange?.(row, col, row, col);
    }, { row: sourceRow, col: sourceCol });
    await page.locator("#formulaInput").click();
    await page.keyboard.press(process.platform === "darwin" ? "Meta+A" : "Control+A");
    await page.keyboard.press("Backspace");
    await page.keyboard.type("=SUM(");
    await page.locator(`.sheet-tab[data-sheet-index="${target.sheetIndex}"]`).click();
    await page.waitForFunction((sheetIndex) => document.querySelector(`.sheet-tab[data-sheet-index="${sheetIndex}"]`)?.classList.contains("is-active"), target.sheetIndex, { timeout: 5000 });
    await page.waitForSelector('.sheet-cell[data-address="A1"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    try {
      await page.waitForFunction(() => {
        const value = document.querySelector("#formulaInput")?.value || "";
        return value.startsWith("=SUM(") && value.includes("!") && value.includes("A1");
      }, null, { timeout: 5000 });
    } catch (error) {
      const debug = await page.evaluate(() => ({
        inputValue: document.querySelector("#formulaInput")?.value || "",
        activeElement: document.activeElement?.id || document.activeElement?.className || document.activeElement?.tagName || "",
        activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
        selectedLabel: document.querySelector("#selectedCellLabel")?.textContent?.trim() || "",
      }));
      throw new Error(`Cross-sheet formula reference was not inserted: ${JSON.stringify(debug)}`);
    }
    const inputValue = await page.locator("#formulaInput").inputValue();
    await page.keyboard.type(")");
    await page.keyboard.press("Enter");
    await page.waitForTimeout(100);
    await page.waitForFunction((sheetIndex) => document.querySelector(`.sheet-tab[data-sheet-index="${sheetIndex}"]`)?.classList.contains("is-active"), source.sheetIndex, { timeout: 5000 });
    const committed = await page.evaluate(({ sheetName, row, col }) => {
      const cell = window.__webSheetDevCellModel?.(sheetName, row, col);
      return cell?.raw || cell?.formula || "";
    }, {
      sheetName: source.name,
      row: sourceRow,
      col: sourceCol,
    });
    return {
      skipped: false,
      source: source.name,
      sourceAddress: crossSheetSourceAddress,
      target: target.name,
      inputValue: `${inputValue})`,
      committed,
      activeSheet: await page.evaluate(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || ""),
      status: await page.evaluate(() => document.querySelector("#statusText")?.textContent?.trim() || ""),
      formulaInputAfterEnter: await page.locator("#formulaInput").inputValue(),
      activeElementAfterEnter: await page.evaluate(() => document.activeElement?.id || document.activeElement?.className || document.activeElement?.tagName || ""),
    };
  })();



  const formulaReferenceInsert = {
    targets: formulaReferenceTargets,
    barClick: formulaReferenceBarClick,
    barDrag: formulaReferenceBarDrag,
    replace: formulaReferenceReplace,
    columnClick: formulaReferenceColumnClick,
    columnDrag: formulaReferenceColumnDrag,
    rowClick: formulaReferenceRowClick,
    rowDrag: formulaReferenceRowDrag,
    columnF4: formulaReferenceColumnF4,
    rowF4: formulaReferenceRowF4,
    cellClick: formulaReferenceCellClick,
    colors: formulaReferenceColors,
    f4: formulaReferenceF4,
    f4Expected: formulaReferenceF4Expected,
    crossSheet: formulaReferenceCrossSheet,
  };



  await page.evaluate(() => {
    const host = document.querySelector("#sheetHost");
    if (!host) return;
    host.scrollTop = 0;
    host.scrollLeft = 0;
    host.dispatchEvent(new Event("scroll"));
  });


  await page.waitForFunction(() => document.querySelector('.sheet-cell[data-address="A2"]'), null, { timeout: 10000 });



  const directTypingAddresses = await page.evaluate(() => {
    const reserved = new Set(["A3", "A4", "A5", "A6", "A7", "B2", "B3", "B14", "D14", "G6", "H6", "I6", "J6"]);
    return [...document.querySelectorAll('.sheet-cell[data-address]')]
      .map((cell) => ({
        address: cell.dataset.address,
        text: cell.querySelector(".cell-content")?.textContent || "",
      }))
      .filter((item) => !reserved.has(item.address) && item.text === "")
      .slice(0, 4)
      .map((item) => item.address);
  });


  if (directTypingAddresses.length < 4) throw new Error("No visible empty cells were available for direct typing, defined name, and link testing.");



  const revealDirectTestCell = async (address) => {
    await page.evaluate(() => {
      const host = document.querySelector("#sheetHost");
      if (!host) return;
      host.scrollTop = 0;
      host.scrollLeft = 0;
      host.dispatchEvent(new Event("scroll"));
    });
    await page.waitForFunction((targetAddress) => document.querySelector(`.sheet-cell[data-address="${targetAddress}"]`), address, { timeout: 10000 });
  };



  const typeCellByDirectKeyboard = async (address, value) => {
    await revealDirectTestCell(address);
    await page.locator(`.sheet-cell[data-address="${address}"]`).click();
    await page.keyboard.type(String(value));
    await page.keyboard.press("Enter");
    await page.waitForFunction(
      ({ address: targetAddress, value: targetValue }) =>
        document.querySelector(`.sheet-cell[data-address="${targetAddress}"]`)?.getAttribute("contenteditable") === null &&
        document.querySelector(`.sheet-cell[data-address="${targetAddress}"] .cell-content`)?.textContent === String(targetValue),
      { address, value },
      { timeout: 10000 },
    );
  };



  await page.locator(`.sheet-cell[data-address="${directTypingAddresses[0]}"]`).click();


  await page.keyboard.press("X");


  await page.waitForFunction(
    (address) =>
      document.querySelector(`.sheet-cell[data-address="${address}"]`)?.getAttribute("contenteditable") === "true" &&
      document.querySelector(`.sheet-cell[data-address="${address}"]`)?.textContent === "X",
    directTypingAddresses[0],
    { timeout: 10000 },
  );


  await page.keyboard.press("Enter");


  await page.waitForFunction(
    (address) =>
      document.querySelector(`.sheet-cell[data-address="${address}"]`)?.getAttribute("contenteditable") === null &&
      document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent === "X",
    directTypingAddresses[0],
    { timeout: 10000 },
  );


  await typeCellByDirectKeyboard(directTypingAddresses[1], "Y");


  const directTyping = await page.evaluate((addresses) => {
    const first = document.querySelector(`.sheet-cell[data-address="${addresses[0]}"]`);
    const second = document.querySelector(`.sheet-cell[data-address="${addresses[1]}"]`);
    return {
      firstAddress: addresses[0],
      secondAddress: addresses[1],
      firstContenteditable: first?.getAttribute("contenteditable"),
      secondContenteditable: second?.getAttribute("contenteditable"),
      firstText: first?.querySelector(".cell-content")?.textContent,
      secondText: second?.querySelector(".cell-content")?.textContent,
      status: document.querySelector("#statusText")?.textContent,
    };
  }, directTypingAddresses);



  const nameLinkAddress = directTypingAddresses[2];


  const formulaNameAddress = directTypingAddresses[3];


  const absoluteNameLinkAddress = (() => {
    const match = nameLinkAddress.match(/^([A-Z]+)(\d+)$/);
    return match ? `$${match[1]}$${match[2]}` : nameLinkAddress;
  })();


  await typeCellByDirectKeyboard(nameLinkAddress, "42");


  await revealDirectTestCell(nameLinkAddress);


  await page.click(`.sheet-cell[data-address="${nameLinkAddress}"]`, { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="define-name"]');


  await page.waitForSelector("#nameDialog");


  const nameDialogInitial = await page.evaluate(() => ({
    title: document.querySelector("#nameDialogTitle")?.textContent,
    nameValue: document.querySelector('[data-name-field="name"]')?.value,
    refersTo: document.querySelector('[data-name-field="refersTo"]')?.value,
    scopeOptions: [...document.querySelectorAll('[data-name-field="scope"] option')].map((option) => option.textContent),
    accessKeys: [...document.querySelectorAll("#nameDialog .menu-access-key")].map((item) => item.textContent),
    accessKeysUnderlined: [...document.querySelectorAll("#nameDialog .menu-access-key")].every((item) =>
      getComputedStyle(item).textDecorationLine.includes("underline"),
    ),
  }));


  await page.locator('[data-name-field="name"]').fill("E2EName");


  await page.click('[data-name-dialog-action="ok"]');


  await page.waitForFunction(() => !document.querySelector("#nameDialog"), null, { timeout: 10000 });


  await revealDirectTestCell(formulaNameAddress);


  await page.locator(`.sheet-cell[data-address="${formulaNameAddress}"]`).click();


  await page.keyboard.type("=E2EName+8");


  await page.keyboard.press("Enter");


  await page.waitForFunction(
    (address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent === "50",
    formulaNameAddress,
    { timeout: 10000 },
  );


  const definedNameFormula = await page.evaluate((address) => ({
    address,
    value: document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent,
    status: document.querySelector("#statusText")?.textContent,
  }), formulaNameAddress);


  await revealDirectTestCell(nameLinkAddress);


  await page.click(`.sheet-cell[data-address="${nameLinkAddress}"]`, { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="define-name"]');


  await page.waitForSelector("#nameDialog");


  await page.locator('[data-name-field="name"]').fill("スキル値");


  await page.click('[data-name-dialog-action="ok"]');


  await page.waitForFunction(() => !document.querySelector("#nameDialog"), null, { timeout: 10000 });


  await page.waitForFunction(() => window.__webSheetDevDefinedNames?.().some((definition) => definition.name === "スキル値"), null, {
    timeout: 10000,
  });


  await revealDirectTestCell(formulaNameAddress);


  await page.locator(`.sheet-cell[data-address="${formulaNameAddress}"]`).click();


  await page.keyboard.type("=スキル値+8");


  await page.keyboard.press("Enter");


  await page.waitForFunction(
    (address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent === "50",
    formulaNameAddress,
    { timeout: 10000 },
  );


  const unicodeDefinedNameFormula = await page.evaluate((address) => {
    const addressPoint = (value) => {
      const match = String(value || "").match(/^([A-Z]+)(\d+)$/);
      if (!match) return null;
      const col = [...match[1]].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
      return { row: Number(match[2]), col };
    };
    const point = addressPoint(address);
    const cell = point ? window.__webSheetDevCellModel?.("CHT", point.row, point.col) : null;
    return {
      address,
      formula: cell?.formula || "",
      raw: cell?.raw ?? null,
      cached: cell?.cached ?? null,
      value: document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent,
      names: window.__webSheetDevDefinedNames?.().map((definition) => definition.name) || [],
      status: document.querySelector("#statusText")?.textContent,
    };
  }, formulaNameAddress);


  await revealDirectTestCell(nameLinkAddress);


  await page.click(`.sheet-cell[data-address="${nameLinkAddress}"]`, { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="define-name"]');


  await page.waitForSelector("#nameDialog");


  await page.locator('[data-name-field="name"]').fill("スキル値");


  await page.click('[data-name-dialog-action="ok"]');


  await page.waitForSelector("#messageBox", { timeout: 5000 });


  const duplicateDefinedName = await page.evaluate(() => ({
    title: document.querySelector("#messageBoxTitle")?.textContent || "",
    message: document.querySelector("#messageBoxText")?.textContent || "",
    dialogStillOpen: Boolean(document.querySelector("#nameDialog")),
    count: window.__webSheetDevDefinedNames?.().filter((definition) => definition.name === "スキル値").length || 0,
  }));


  await page.click("[data-message-box-close]");


  await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });


  await page.click('[data-name-dialog-action="cancel"]');


  await page.waitForFunction(() => !document.querySelector("#nameDialog"), null, { timeout: 10000 });


  await revealDirectTestCell(nameLinkAddress);


  await page.click(`.sheet-cell[data-address="${nameLinkAddress}"]`, { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="define-name"]');


  await page.waitForSelector("#nameDialog");


  await page.locator('[data-name-field="name"]').fill("ローカル値");


  await page.selectOption('[data-name-field="scope"]', "sheet:0");


  await page.locator('[data-name-field="refersTo"]').fill(`=${absoluteNameLinkAddress}`);


  await page.click('[data-name-dialog-action="ok"]');


  await page.waitForFunction(() => !document.querySelector("#nameDialog"), null, { timeout: 10000 });


  await page.waitForFunction(
    () => window.__webSheetDevDefinedNames?.().some((definition) => definition.name === "ローカル値" && definition.scope === "sheet"),
    null,
    { timeout: 10000 },
  );


  await revealDirectTestCell(formulaNameAddress);


  await page.locator(`.sheet-cell[data-address="${formulaNameAddress}"]`).click();


  await page.keyboard.type("='CHT'!ローカル値+9");


  await page.keyboard.press("Enter");


  await page.waitForFunction(
    (address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent === "51",
    formulaNameAddress,
    { timeout: 10000 },
  );


  const localDefinedNameFormula = await page.evaluate((address) => {
    const addressPoint = (value) => {
      const match = String(value || "").match(/^([A-Z]+)(\d+)$/);
      if (!match) return null;
      const col = [...match[1]].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
      return { row: Number(match[2]), col };
    };
    const point = addressPoint(address);
    const cell = point ? window.__webSheetDevCellModel?.("CHT", point.row, point.col) : null;
    return {
      address,
      formula: cell?.formula || "",
      value: document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent,
      localName: window.__webSheetDevDefinedNames?.().find((definition) => definition.name === "ローカル値") || null,
      status: document.querySelector("#statusText")?.textContent,
    };
  }, formulaNameAddress);


  await revealDirectTestCell(nameLinkAddress);


  await page.click(`.sheet-cell[data-address="${nameLinkAddress}"]`, { button: "right" });


  await page.waitForSelector("#cellContextMenu");


  await page.click('#cellContextMenu [data-cell-menu-action="link"]');


  await page.waitForSelector("#linkDialog");
  Object.assign(scenarioState, { context, page, messages, result, runtimeSource, saveAsDialog, expectedReloadWorkbookName, standaloneBlockedRequests, standaloneOffline, standaloneRemoteFallback, standaloneEmptyModel, standaloneMalformedModel, standaloneWrappedModel, launcherBlockedRequests, standaloneLauncher, ar2Gridline, imageAndFont, imageSelection, footerLayout, realCtrlWheelZoom, zoomPopup, footerBeforeScroll, tabScroll, singleClickEdit, doubleClickEdit, editAfterDeferredRender, editAlignment, editClickCaret, editTextSelection, editNoEllipsis, cellEditAltEnter, cellEditAltEnterDuringIme, editDragOutsideSelection, cellEditImeEnter, editScrollbar, editExitRender, multilineCommitDisplay, formulaEmptyAddress, formulaBarInitial, formulaBarSizing, formulaBarDuringEdit, formulaBarEscape, formulaBarEnter, formulaAutocomplete, formulaCustomFunctions, formulaReferenceInsert, revealDirectTestCell, typeCellByDirectKeyboard, directTyping, nameLinkAddress, formulaNameAddress, absoluteNameLinkAddress, nameDialogInitial, definedNameFormula, unicodeDefinedNameFormula, duplicateDefinedName, localDefinedNameFormula });
}

export {
  runObjectsAndFormulas,
};
