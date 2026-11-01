// Sequential E2E feature module.
import {
  JSZip,
  appUrl,
  workbookPath,
} from "../../e2e-real-upload.mjs";
import {
  createAutoFilterMenuWorkbook,
  createConsolidateWorkbook,
  createDataAnalysisWorkbook,
  createForecastWorkbook,
  createOutlineWorkbook,
  createRelationshipsWorkbook,
  createRemoveDuplicatesWorkbook,
  createSortBlankWorkbook,
  createSortCaseWorkbook,
  createSortColumnsWorkbook,
  createSortDialogWorkbook,
  createSubtotalWorkbook,
  createTextToColumnsAdvancedWorkbook,
  createTextToColumnsFixedWorkbook,
  createTextToColumnsWorkbook,
  exportedOutlineWorkbookState,
  waitForUploadedWorkbook,
} from "./fixtures-and-model.mjs";
import {
  createSortCustomListWorkbook,
  createSortDateWorkbook,
  createSortFormulaWorkbook,
} from "./editing-and-validation.mjs";


async function exerciseDataGetConnectionMenus(context) {
  const workbookPath = await createAutoFilterMenuWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?dataGetConnections=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="J1"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="J1"]').click();
    await page.evaluate(() => {
      window.__webSheetDevSelectRange?.(1, 10, 1, 10);
    });
    await page.locator('[data-ribbon-tab="data"]').click();

    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    const getMenu = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#dataCommandMenu button")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
      accessKeys: [...document.querySelectorAll("#dataCommandMenu button")].map((button) => button.dataset.accessKey || ""),
      disabled: [...document.querySelectorAll("#dataCommandMenu button:disabled")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
      focusedAction: document.activeElement?.dataset?.dataMenuAction || "",
      j1BeforeShortcut: window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 1, 10)?.raw || "",
    }));
    await page.keyboard.press("W");
    await page.waitForSelector('[data-data-dialog="web-import"]', { timeout: 5000 });
    const dataGetShortcut = await page.evaluate(() => ({
      j1AfterShortcut: window.__webSheetDevCellModel?.(window.__webSheetDevWorkbookState?.().activeSheetName || "", 1, 10)?.raw || "",
      editingCellCount: document.querySelectorAll(".sheet-cell[contenteditable='true']").length,
      focusedDialogField: document.activeElement?.matches?.("[data-data-field='sourceUrl']") || false,
      dialogTitle: document.querySelector("#webImportTitle")?.textContent || "",
      menuClosed: !document.querySelector("#dataCommandMenu"),
    }));
    const webDialog = await page.evaluate(() => ({
      title: document.querySelector("#webImportTitle")?.textContent || "",
      captions: [...document.querySelectorAll("[data-data-dialog='web-import'] .excel-dialog-caption")].map((item) => item.textContent.trim()),
      note: document.querySelector("[data-data-dialog='web-import'] .data-dialog-note")?.textContent || "",
    }));
    const webDataUrl = "data:text/csv;charset=utf-8,Name%2CScore%0AWeb%2C9";
    await page.locator("[data-data-field='sourceUrl']").fill(webDataUrl);
    await page.locator('[data-data-dialog-action="load"]').click();
    await page.waitForSelector('[data-data-dialog="web-import"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="J1"]')?.textContent?.includes("Name") &&
        document.querySelector('.sheet-cell[data-address="K2"]')?.textContent?.includes("9"),
      null,
      { timeout: 5000 },
    );
    const webImport = await page.evaluate(() => ({
      j1: document.querySelector('.sheet-cell[data-address="J1"]')?.textContent.trim() || "",
      k2: document.querySelector('.sheet-cell[data-address="K2"]')?.textContent.trim() || "",
      status: document.querySelector("#statusText")?.textContent || "",
      query: window.__webSheetDevQueries?.()[0]?.query || null,
    }));

    await page.locator('[data-ribbon-command="data-recent-sources"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    const recentMenu = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#dataCommandMenu button")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    }));
    await page.locator('#dataCommandMenu [data-data-menu-action="recent-source"]').first().click();
    await page.waitForFunction(
      () => document.querySelector("#statusText")?.textContent?.includes("WebQuery1 の範囲へ移動"),
      null,
      { timeout: 5000 },
    );
    const recentSelect = await page.evaluate(() => ({
      selected: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
      frameStyle: document.querySelector(".sheet-selection-frame")?.getAttribute("style") || "",
    }));

    await page.locator('.sheet-cell[data-address="M1"]').click();
    await page.locator('[data-ribbon-command="data-existing-connections"]').click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { timeout: 5000 });
    const existingDialog = await page.evaluate(() => ({
      title: document.querySelector("#existingConnectionsTitle")?.textContent || "",
      options: [...document.querySelectorAll("[data-existing-connection-list] option")].map((option) => option.textContent.trim()),
      details: document.querySelector("[data-existing-connection-details]")?.textContent || "",
      actions: [...document.querySelectorAll("[data-data-dialog='existing-connections'] [data-data-dialog-action]")].map((button) => ({
        text: button.textContent.trim(),
        disabled: button.disabled,
      })),
    }));
    await page.locator('[data-data-dialog-action="existing-open"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    const importDataDialog = await page.evaluate(() => ({
      title: document.querySelector("#importDataTitle")?.textContent || "",
      note: document.querySelector("[data-data-dialog='import-data'] .data-dialog-note")?.textContent || "",
      modes: [...document.querySelectorAll("[data-data-dialog='import-data'] input[name='importDataMode']")].map((input) => ({
        value: input.value,
        label: input.closest("label")?.textContent.replace(/\s+/g, " ").trim() || "",
        disabled: input.disabled,
        checked: input.checked,
      })),
      locations: [...document.querySelectorAll("[data-data-dialog='import-data'] input[name='importDataLocation']")].map((input) => ({
        value: input.value,
        label: input.closest("label")?.textContent.replace(/\s+/g, " ").trim() || "",
        disabled: input.disabled,
        checked: input.checked,
      })),
      destination: document.querySelector("[data-data-field='destination']")?.value || "",
      buttons: [...document.querySelectorAll("[data-data-dialog='import-data'] [data-data-dialog-action]")].map((button) => button.textContent.trim()),
    }));
    await page.locator("[data-data-dialog='import-data'] input[name='importDataMode'][value='connectionOnly']").check();
    await page.locator("[data-data-field='addToDataModel']").check();
    const connectionOnlySync = await page.evaluate(() => ({
      destinationDisabled: Boolean(document.querySelector("[data-data-field='destination']")?.disabled),
      locationsDisabled: [...document.querySelectorAll("[data-data-dialog='import-data'] input[name='importDataLocation']")].every((input) => input.disabled),
      addToDataModel: Boolean(document.querySelector("[data-data-field='addToDataModel']")?.checked),
    }));
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => window.__webSheetDevQueries?.()[0]?.query?.loadMode === "connectionOnly" &&
        window.__webSheetDevQueries?.()[0]?.query?.addToDataModel === true,
      null,
      { timeout: 5000 },
    );
    const connectionOnlyImport = await page.evaluate(() => ({
      status: document.querySelector("#statusText")?.textContent || "",
      query: window.__webSheetDevQueries?.()[0]?.query || null,
      dataModel: window.__webSheetDevDataModel?.() || null,
      m1: document.querySelector('.sheet-cell[data-address="M1"]')?.textContent.trim() || "",
    }));

    await page.locator('[data-ribbon-command="data-manage-data-model"]').click();
    await page.waitForSelector('[data-data-dialog="data-model"]', { timeout: 5000 });
    const dataModelDialog = await page.evaluate(() => ({
      title: document.querySelector("#dataModelTitle")?.textContent || "",
      options: [...document.querySelectorAll("[data-data-model-table-list] option")].map((option) => option.textContent.trim()),
      details: document.querySelector("[data-data-model-details]")?.textContent || "",
      relationships: document.querySelector("[data-data-model-relationships]")?.textContent || "",
      actions: [...document.querySelectorAll("[data-data-dialog='data-model'] [data-data-dialog-action]")].map((button) => ({
        text: button.textContent.trim(),
        disabled: button.disabled,
      })),
    }));
    await page.locator("[data-data-dialog='data-model'] [data-data-dialog-action='cancel']").click();
    await page.waitForSelector('[data-data-dialog="data-model"]', { state: "detached", timeout: 5000 });

    await page.locator('[data-ribbon-command="data-existing-connections"]').click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="existing-open"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-dialog='import-data'] input[name='importDataMode'][value='table']").check();
    await page.waitForFunction(
      () => document.querySelector("[data-data-field='destination']")?.value?.includes("M1"),
      null,
      { timeout: 5000 },
    );
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="M1"]')?.textContent?.includes("Name") &&
        document.querySelector('.sheet-cell[data-address="N2"]')?.textContent?.includes("9"),
      null,
      { timeout: 5000 },
    );
    const importedConnection = await page.evaluate(() => ({
      m1: document.querySelector('.sheet-cell[data-address="M1"]')?.textContent.trim() || "",
      n2: document.querySelector('.sheet-cell[data-address="N2"]')?.textContent.trim() || "",
      selected: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
      query: window.__webSheetDevQueries?.()[0]?.query || null,
    }));

    await page.locator('[data-ribbon-command="data-existing-connections"]').click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="existing-open"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-dialog='import-data'] input[name='importDataMode'][value='pivotTable']").check();
    await page.locator("[data-data-dialog='import-data'] input[name='importDataLocation'][value='new']").check();
    const pivotImportDialog = await page.evaluate(() => ({
      destinationDisabled: Boolean(document.querySelector("[data-data-field='destination']")?.disabled),
      pivotChecked: Boolean(document.querySelector("[data-data-dialog='import-data'] input[name='importDataMode'][value='pivotTable']")?.checked),
      newWorksheetChecked: Boolean(document.querySelector("[data-data-dialog='import-data'] input[name='importDataLocation'][value='new']")?.checked),
    }));
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.includes("PivotTable") &&
        document.querySelector('.sheet-cell[data-address="A1"]')?.textContent?.trim(),
      null,
      { timeout: 5000 },
    );
    const pivotImport = await page.evaluate(() => ({
      active: document.querySelector(".sheet-tab.is-active")?.textContent.trim() || "",
      headers: ["A1", "B1", "C1"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"]`)?.textContent.trim() || ""),
      tableCount: window.__webSheetDevQueries?.().length || 0,
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-existing-connections"]').click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="existing-open"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-dialog='import-data'] input[name='importDataMode'][value='pivotChart']").check();
    await page.locator("[data-data-dialog='import-data'] input[name='importDataLocation'][value='new']").check();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.includes("PivotChart") &&
        document.querySelectorAll(".sheet-image").length > 0,
      null,
      { timeout: 5000 },
    );
    const pivotChartImport = await page.evaluate(() => ({
      active: document.querySelector(".sheet-tab.is-active")?.textContent.trim() || "",
      imageCount: document.querySelectorAll(".sheet-image").length,
      chartName: document.querySelector(".sheet-image")?.getAttribute("alt") || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('[data-ribbon-command="data-existing-connections"]').click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="existing-properties"]').click();
    await page.waitForSelector('[data-data-dialog="query-properties"]', { timeout: 5000 });
    await page.locator('[data-query-properties-tab="definition"]').click();
    const existingProperties = await page.evaluate(() => ({
      connectionString: document.querySelector("[data-data-field='connectionString']")?.value || "",
      commandText: document.querySelector("[data-data-field='commandText']")?.value || "",
      typeNote: document.querySelector("[data-query-properties-tab-panel='usage'] .data-dialog-note")?.textContent || "",
    }));
    await page.locator('[data-data-dialog-action="cancel"]').click();
    await page.waitForSelector('[data-data-dialog="query-properties"]', { state: "detached", timeout: 5000 });
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { timeout: 5000 });
    const existingAfterPropertiesCancel = await page.evaluate(() => ({
      title: document.querySelector("#existingConnectionsTitle")?.textContent || "",
      options: [...document.querySelectorAll("[data-existing-connection-list] option")].map((option) => option.textContent.trim()),
    }));
    await page.locator('[data-data-dialog-action="existing-delete"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const existingDeletePrompt = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("#messageBoxOverlay [data-message-box-close]")].map((button) => button.textContent.trim()),
      dialogVisible: Boolean(document.querySelector('[data-data-dialog="existing-connections"]')),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-result='cancel']").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    const existingDeleteCancelCount = await page.locator("[data-existing-connection-list] option").count();
    await page.locator("[data-data-dialog='existing-connections'] [data-data-dialog-action='cancel']").click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { state: "detached", timeout: 5000 });

    await page.locator('[data-ribbon-command="data-refresh-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    const refreshMenu = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#dataCommandMenu button")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
      disabled: [...document.querySelectorAll("#dataCommandMenu button:disabled")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    }));
    await page.locator('#dataCommandMenu [data-data-menu-action="refresh-all"]').click();
    await page.waitForFunction(
      () => document.querySelector("#statusText")?.textContent?.includes("1 件の接続を更新"),
      null,
      { timeout: 5000 },
    );
    const refreshAll = await page.evaluate(() => ({
      status: document.querySelector("#statusText")?.textContent || "",
      query: window.__webSheetDevQueries?.()[0]?.query || null,
    }));

    await page.locator('[data-ribbon-command="data-queries-connections"]').click();
    await page.waitForSelector("#queryPane", { timeout: 5000 });
    const webQueryPaneBeforeEdit = await page.evaluate(() => ({
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      editButtons: [...document.querySelectorAll("[data-query-pane-action='edit']")].length,
    }));
    const webQueryEditClicked = await page.evaluate(() => {
      const row = [...document.querySelectorAll(".query-pane-item")].find((item) => item.textContent.includes("WebQuery1"));
      const button = row?.querySelector("[data-query-pane-action='edit']");
      if (!button) return false;
      button.click();
      return true;
    });
    if (!webQueryEditClicked) throw new Error("WebQuery1 edit button was not found in the Queries & Connections pane.");
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    const webQueryEditDialog = await page.evaluate(() => ({
      title: document.querySelector("#queryTransformTitle")?.textContent || "",
      queryName: document.querySelector("[data-data-field='queryName']")?.value || "",
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, " ").trim() || "",
      loadToVisible: Boolean(document.querySelector("[data-data-dialog-action='query-transform-load-to']")),
    }));
    await page.locator("[data-data-dialog='query-transform-editor'] [data-data-dialog-action='cancel']").click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { state: "detached", timeout: 5000 });

    const openQueryPaneContextMenu = async (queryName) => {
      const opened = await page.evaluate((name) => {
        const row = [...document.querySelectorAll(".query-pane-item")].find((item) =>
          item.textContent.includes(name),
        );
        if (!row) return false;
        const rect = row.getBoundingClientRect();
        row.dispatchEvent(new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + Math.min(24, Math.max(4, rect.width / 2)),
          clientY: rect.top + Math.min(18, Math.max(4, rect.height / 2)),
          button: 2,
        }));
        return true;
      }, queryName);
      if (!opened) throw new Error(`Query pane row was not found for context menu: ${queryName}`);
      await page.waitForSelector("#queryPaneContextMenu", { timeout: 5000 });
    };

    await openQueryPaneContextMenu("WebQuery1");
    const webQueryContextMenu = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#queryPaneContextMenu [data-query-pane-menu-action]")].map((button) =>
        button.textContent.replace(/\s+/g, " ").trim(),
      ),
      disabled: [...document.querySelectorAll("#queryPaneContextMenu [data-query-pane-menu-action]:disabled")].map((button) =>
        button.textContent.replace(/\s+/g, " ").trim(),
      ),
      focusedBefore: document.activeElement?.dataset?.queryPaneMenuAction || "",
    }));
    await page.keyboard.press("ArrowDown");
    Object.assign(webQueryContextMenu, await page.evaluate(() => ({
      focusedAfterDown: document.activeElement?.dataset?.queryPaneMenuAction || "",
    })));
    await page.locator("#queryPaneContextMenu [data-query-pane-menu-action='duplicate']").click();
    await page.waitForFunction(
      () => (window.__webSheetDevQueries?.() || []).some((entry) => entry.query?.name === "WebQuery1 (2)" && entry.query?.loadMode === "connectionOnly"),
      null,
      { timeout: 5000 },
    );
    const duplicatedQuery = await page.evaluate(() => {
      const entry = (window.__webSheetDevQueries?.() || []).find((item) => item.query?.name === "WebQuery1 (2)");
      return {
        query: entry?.query || null,
        rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await openQueryPaneContextMenu("WebQuery1");
    await page.locator("#queryPaneContextMenu [data-query-pane-menu-action='reference']").click();
    await page.waitForFunction(
      () => (window.__webSheetDevQueries?.() || []).some((entry) =>
        entry.query?.sourceType === "queryReference" && entry.query?.sourceQueryName === "WebQuery1",
      ),
      null,
      { timeout: 5000 },
    );
    const referenceCreated = await page.evaluate(() => {
      const entry = (window.__webSheetDevQueries?.() || []).find((item) =>
        item.query?.sourceType === "queryReference" && item.query?.sourceQueryName === "WebQuery1",
      );
      return {
        name: entry?.query?.name || "",
        query: entry?.query || null,
        rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await openQueryPaneContextMenu(referenceCreated.name);
    await page.locator("#queryPaneContextMenu [data-query-pane-menu-action='rename']").click();
    await page.waitForSelector('[data-data-dialog="query-rename"]', { timeout: 5000 });
    const renameDialog = await page.evaluate(() => ({
      title: document.querySelector("#queryRenameTitle")?.textContent || "",
      value: document.querySelector("[data-data-field='queryRenameName']")?.value || "",
    }));
    await page.locator("[data-data-field='queryRenameName']").fill("WebQueryRef");
    await page.locator("[data-data-dialog='query-rename'] [data-data-dialog-action='ok']").click();
    await page.waitForSelector('[data-data-dialog="query-rename"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => (window.__webSheetDevQueries?.() || []).some((entry) =>
        entry.query?.name === "WebQueryRef" && entry.query?.sourceType === "queryReference",
      ),
      null,
      { timeout: 5000 },
    );
    const renamedReference = await page.evaluate(() => {
      const entry = (window.__webSheetDevQueries?.() || []).find((item) => item.query?.name === "WebQueryRef");
      return {
        query: entry?.query || null,
        rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await openQueryPaneContextMenu("WebQueryRef");
    await page.locator("#queryPaneContextMenu [data-query-pane-menu-action='properties']").click();
    await page.waitForSelector('[data-data-dialog="query-properties"]', { timeout: 5000 });
    await page.locator('[data-query-properties-tab="definition"]').click();
    const referenceProperties = await page.evaluate(() => ({
      title: document.querySelector("#queryPropertiesTitle")?.textContent || "",
      name: document.querySelector("[data-data-field='name']")?.value || "",
      typeNote: document.querySelector("[data-query-properties-tab-panel='usage'] .data-dialog-note")?.textContent || "",
      connectionString: document.querySelector("[data-data-field='connectionString']")?.value || "",
      commandText: document.querySelector("[data-data-field='commandText']")?.value || "",
      commandType: document.querySelector("[data-data-field='commandType']")?.value || "",
    }));
    await page.locator("[data-data-dialog='query-properties'] [data-data-dialog-action='cancel']").click();
    await page.waitForSelector('[data-data-dialog="query-properties"]', { state: "detached", timeout: 5000 });

    await openQueryPaneContextMenu("WebQueryRef");
    await page.locator("#queryPaneContextMenu [data-query-pane-menu-action='load-to']").click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-dialog='import-data'] input[name='importDataMode'][value='table']").check();
    await page.locator("[data-data-field='destination']").fill("Filter!P1");
    await page.locator("[data-data-dialog='import-data'] [data-data-dialog-action='ok']").click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="P1"]')?.textContent?.includes("Name") &&
        document.querySelector('.sheet-cell[data-address="Q2"]')?.textContent?.includes("9"),
      null,
      { timeout: 5000 },
    );
    const referenceLoadTo = await page.evaluate(() => {
      const entry = (window.__webSheetDevQueries?.() || []).find((item) => item.query?.name === "WebQueryRef");
      return {
        p1: document.querySelector('.sheet-cell[data-address="P1"]')?.textContent.trim() || "",
        q2: document.querySelector('.sheet-cell[data-address="Q2"]')?.textContent.trim() || "",
        query: entry?.query || null,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    await openQueryPaneContextMenu("WebQueryRef");
    await page.locator("#queryPaneContextMenu [data-query-pane-menu-action='refresh']").click();
    await page.waitForFunction(
      () => document.querySelector("#statusText")?.textContent?.includes("WebQueryRef を更新"),
      null,
      { timeout: 5000 },
    );
    const referenceRefresh = await page.evaluate(() => {
      const entry = (window.__webSheetDevQueries?.() || []).find((item) => item.query?.name === "WebQueryRef");
      return {
        p1: document.querySelector('.sheet-cell[data-address="P1"]')?.textContent.trim() || "",
        q2: document.querySelector('.sheet-cell[data-address="Q2"]')?.textContent.trim() || "",
        refreshCount: entry?.query?.refreshCount || 0,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="other-data-sources"]').click();
    await page.waitForSelector("#dataCommandMenu.is-other-data-sources", { timeout: 5000 });
    const blankQueryMenu = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("#dataCommandMenu button")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
      blankDisabled: Boolean(document.querySelector('#dataCommandMenu [data-data-menu-action="get-blank-query"]')?.disabled),
    }));
    await page.locator('#dataCommandMenu [data-data-menu-action="get-blank-query"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    const blankQueryEditor = await page.evaluate(() => ({
      title: document.querySelector("#queryTransformTitle")?.textContent || "",
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      formula: document.querySelector(".power-query-formula-bar")?.textContent || "",
      queryName: document.querySelector("[data-data-field='queryName']")?.value || "",
      loadToVisible: Boolean(document.querySelector("[data-data-dialog-action='query-transform-load-to']")),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    const blankQueryImportDialog = await page.evaluate(() => ({
      title: document.querySelector("#importDataTitle")?.textContent || "",
      note: document.querySelector("[data-data-dialog='import-data'] .data-dialog-note")?.textContent || "",
      modes: [...document.querySelectorAll("[data-data-dialog='import-data'] input[name='importDataMode']")].map((input) => ({
        value: input.value,
        label: input.closest("label")?.textContent.replace(/\s+/g, " ").trim() || "",
        disabled: input.disabled,
        checked: input.checked,
      })),
    }));
    await page.locator("[data-data-dialog='import-data'] input[name='importDataMode'][value='connectionOnly']").check();
    await page.locator("[data-data-dialog='import-data'] [data-data-dialog-action='ok']").click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => (window.__webSheetDevQueries?.() || []).some((entry) =>
        entry.query?.sourceType === "blankQuery" && entry.query?.loadMode === "connectionOnly",
      ),
      null,
      { timeout: 5000 },
    );
    const blankQueryConnection = await page.evaluate(() => {
      const entry = (window.__webSheetDevQueries?.() || []).find((item) => item.query?.sourceType === "blankQuery");
      return {
        query: entry?.query || null,
        sheetIndex: entry?.sheetIndex ?? null,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-command="data-existing-connections"]').click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { timeout: 5000 });
    const blankQueryExistingSelected = await page.evaluate(() => {
      const entry = (window.__webSheetDevQueries?.() || []).find((item) => item.query?.sourceType === "blankQuery");
      const select = document.querySelector("[data-existing-connection-list]");
      if (!entry || !select) return false;
      select.value = `${entry.sheetIndex}::${entry.query.id}`;
      select.dispatchEvent(new Event("change", { bubbles: true }));
      return true;
    });
    const blankQueryExistingDetails = await page.evaluate(() => ({
      selected: document.querySelector("[data-existing-connection-list]")?.selectedOptions?.[0]?.textContent.trim() || "",
      details: document.querySelector("[data-existing-connection-details]")?.textContent || "",
      canOpen: !document.querySelector("[data-data-dialog-action='existing-open']")?.disabled,
      canProperties: !document.querySelector("[data-data-dialog-action='existing-properties']")?.disabled,
    }));
    await page.locator("[data-data-dialog='existing-connections'] [data-data-dialog-action='cancel']").click();
    await page.waitForSelector('[data-data-dialog="existing-connections"]', { state: "detached", timeout: 5000 });

    await page.locator('[data-ribbon-command="data-queries-connections"]').click();
    await page.waitForSelector("#queryPane", { timeout: 5000 });
    const blankQueryPane = await page.evaluate(() => ({
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      tabs: [...document.querySelectorAll(".query-pane-tabs button")].map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    }));
    const blankPanePropertiesClicked = await page.evaluate(() => {
      const row = [...document.querySelectorAll(".query-pane-item")].find((item) =>
        item.textContent.includes("空のクエリ接続") || item.textContent.includes("空のクエリ"),
      );
      const button = row?.querySelector("[data-query-pane-action='properties']");
      if (!button) return false;
      button.click();
      return true;
    });
    if (!blankPanePropertiesClicked) throw new Error("Blank query properties button was not found in the Queries & Connections pane.");
    await page.waitForSelector('[data-data-dialog="query-properties"]', { timeout: 5000 });
    await page.locator('[data-query-properties-tab="definition"]').click();
    const blankQueryPaneProperties = await page.evaluate(() => ({
      title: document.querySelector("#queryPropertiesTitle")?.textContent || "",
      name: document.querySelector("[data-data-field='name']")?.value || "",
      typeNote: document.querySelector("[data-query-properties-tab-panel='usage'] .data-dialog-note")?.textContent || "",
      connectionString: document.querySelector("[data-data-field='connectionString']")?.value || "",
      commandType: document.querySelector("[data-data-field='commandType']")?.value || "",
      commandText: document.querySelector("[data-data-field='commandText']")?.value || "",
    }));
    await page.locator('[data-data-dialog-action="cancel"]').click();
    await page.waitForSelector('[data-data-dialog="query-properties"]', { state: "detached", timeout: 5000 });
    const blankLoadToClicked = await page.evaluate(() => {
      const row = [...document.querySelectorAll(".query-pane-item")].find((item) =>
        item.textContent.includes("空のクエリ接続") || item.textContent.includes("空のクエリ"),
      );
      const button = row?.querySelector("[data-query-pane-action='load-to']");
      if (!button) return false;
      button.click();
      return true;
    });
    if (!blankLoadToClicked) throw new Error("Blank query load-to button was not found in the Queries & Connections pane.");
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    const blankQueryLoadToInitial = await page.evaluate(() => ({
      connectionOnlyChecked: Boolean(document.querySelector("[data-data-dialog='import-data'] input[name='importDataMode'][value='connectionOnly']")?.checked),
      destinationDisabled: Boolean(document.querySelector("[data-data-field='destination']")?.disabled),
      locationsDisabled: [...document.querySelectorAll("[data-data-dialog='import-data'] input[name='importDataLocation']")].every((input) => input.disabled),
    }));
    await page.locator("[data-data-dialog='import-data'] [data-data-dialog-action='cancel']").click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });

    return {
      getMenu,
      dataGetShortcut,
      webDialog,
      webImport,
      recentMenu,
      recentSelect,
      existingDialog,
      importDataDialog,
      connectionOnlySync,
      connectionOnlyImport,
      dataModelDialog,
      importedConnection,
      pivotImportDialog,
      pivotImport,
      pivotChartImport,
      existingProperties,
      existingAfterPropertiesCancel,
      existingDeletePrompt,
      existingDeleteCancelCount,
      refreshMenu,
      refreshAll,
      webQueryPaneBeforeEdit,
      webQueryEditDialog,
      webQueryContextMenu,
      duplicatedQuery,
      referenceCreated,
      renameDialog,
      renamedReference,
      referenceProperties,
      referenceLoadTo,
      referenceRefresh,
      blankQueryMenu,
      blankQueryEditor,
      blankQueryImportDialog,
      blankQueryConnection,
      blankQueryExistingSelected,
      blankQueryExistingDetails,
      blankQueryPane,
      blankQueryPaneProperties,
      blankQueryLoadToInitial,
    };
  } finally {
    await page.close();
  }
}



async function exerciseRelationshipsDialog(context) {
  const workbookPath = await createRelationshipsWorkbook();
  const page = await context.newPage();
  const createTableFromCell = async (address, expectedRange) => {
    await page.locator(`.sheet-cell[data-address="${address}"]`).click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.waitForFunction(
      (range) => document.querySelector("[data-data-field='sourceRange']")?.value === range,
      expectedRange,
      { timeout: 5000 },
    );
    await page.locator('[data-data-dialog-action="load"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { state: "detached", timeout: 5000 });
  };
  const managerState = async () =>
    page.evaluate(() => ({
      title: document.querySelector("#relationshipsTitle")?.textContent || "",
      options: [...document.querySelectorAll("[data-relationship-list] option")].map((option) => option.textContent.trim()),
      details: document.querySelector("[data-relationship-details]")?.textContent || "",
      editDisabled: Boolean(document.querySelector("[data-data-dialog-action='relationship-edit']")?.disabled),
      deleteDisabled: Boolean(document.querySelector("[data-data-dialog-action='relationship-delete']")?.disabled),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
  try {
    await page.goto(`${appUrl}?relationships=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="F5"]', { timeout: 10000 });

    await createTableFromCell("A1", "A1:B4");
    await createTableFromCell("D1", "D1:F5");

    await page.locator('[data-ribbon-command="data-relationships"]').click();
    await page.waitForSelector('[data-data-dialog="relationships"]', { timeout: 5000 });
    const empty = await managerState();

    await page.locator('[data-data-dialog-action="relationship-new"]').click();
    await page.waitForSelector('[data-data-dialog="relationship-edit"]', { timeout: 5000 });
    const editInitial = await page.evaluate(() => ({
      title: document.querySelector("#relationshipEditTitle")?.textContent || "",
      sourceTables: [...document.querySelectorAll("[data-data-field='sourceTableKey'] option")].map((option) => option.textContent.trim()),
      sourceColumns: [...document.querySelectorAll("[data-data-field='sourceColumn'] option")].map((option) => option.textContent.trim()),
      targetTables: [...document.querySelectorAll("[data-data-field='targetTableKey'] option")].map((option) => option.textContent.trim()),
      targetColumns: [...document.querySelectorAll("[data-data-field='targetColumn'] option")].map((option) => option.textContent.trim()),
      note: document.querySelector("[data-relationship-edit-note]")?.textContent || "",
    }));
    await page.locator("[data-data-field='sourceColumn']").selectOption("CustomerID");
    await page.locator("[data-data-field='targetColumn']").selectOption("CustomerID");
    const noteAfterColumnChange = await page.locator("[data-relationship-edit-note]").innerText();
    await page.locator('[data-data-dialog-action="relationship-ok"]').click();
    await page.waitForSelector('[data-data-dialog="relationships"]', { timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector("[data-relationship-details]")?.textContent?.includes("Table2") &&
        document.querySelector("[data-relationship-details]")?.textContent?.includes("CustomerID"),
      null,
      { timeout: 5000 },
    );
    const created = await managerState();

    await page.locator('[data-data-dialog-action="relationship-new"]').click();
    await page.waitForSelector('[data-data-dialog="relationship-edit"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceColumn']").selectOption("CustomerID");
    await page.locator("[data-data-field='targetColumn']").selectOption("CustomerID");
    await page.locator('[data-data-dialog-action="relationship-ok"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const duplicatePrompt = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("#messageBoxOverlay [data-message-box-close]")].map((button) => button.textContent.trim()),
      editDialogVisible: Boolean(document.querySelector("[data-data-dialog='relationship-edit']")),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-result='ok']").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await page.locator('[data-data-dialog-action="relationship-cancel"]').click();
    await page.waitForSelector('[data-data-dialog="relationships"]', { timeout: 5000 });

    await page.locator('[data-data-dialog-action="relationship-delete"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const deletePrompt = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("#messageBoxOverlay [data-message-box-close]")].map((button) => button.textContent.trim()),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-result='delete']").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector("[data-relationship-details]")?.textContent?.includes("ありません"),
      null,
      { timeout: 5000 },
    );
    const afterDelete = await managerState();

    await page.locator("[data-data-dialog='relationships'] [data-data-dialog-action='cancel']").click();
    await page.waitForSelector('[data-data-dialog="relationships"]', { state: "detached", timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.locator('[data-ribbon-command="data-relationships"]').click();
    await page.waitForSelector('[data-data-dialog="relationships"]', { timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector("[data-relationship-details]")?.textContent?.includes("Table2"),
      null,
      { timeout: 5000 },
    );
    const undo = await managerState();

    await page.locator("[data-data-dialog='relationships'] [data-data-dialog-action='cancel']").click();
    await page.waitForSelector('[data-data-dialog="relationships"]', { state: "detached", timeout: 5000 });
    await page.keyboard.press("Control+Y");
    await page.locator('[data-ribbon-command="data-relationships"]').click();
    await page.waitForSelector('[data-data-dialog="relationships"]', { timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector("[data-relationship-details]")?.textContent?.includes("ありません"),
      null,
      { timeout: 5000 },
    );
    const redo = await managerState();

    await page.locator('[data-data-dialog-action="relationship-auto-detect"]').click();
    await page.waitForFunction(
      () => document.querySelector("[data-relationship-details]")?.textContent?.includes("Table2") &&
        document.querySelector("#statusText")?.textContent?.includes("検出"),
      null,
      { timeout: 5000 },
    );
    const autoDetected = await managerState();

    await page.locator("[data-data-dialog='relationships'] [data-data-dialog-action='cancel']").click();
    await page.waitForSelector('[data-data-dialog="relationships"]', { state: "detached", timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.locator('[data-ribbon-command="data-relationships"]').click();
    await page.waitForSelector('[data-data-dialog="relationships"]', { timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector("[data-relationship-details]")?.textContent?.includes("ありません"),
      null,
      { timeout: 5000 },
    );
    const autoDetectUndo = await managerState();

    return { empty, editInitial, noteAfterColumnChange, created, duplicatePrompt, deletePrompt, afterDelete, undo, redo, autoDetected, autoDetectUndo };
  } finally {
    await page.close();
  }
}



async function exerciseRefreshAllUndo(context) {
  const workbookPath = await createAutoFilterMenuWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?refreshAllUndo=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="B3"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="load"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { state: "detached", timeout: 5000 });
    await page.locator('[data-ribbon-command="data-queries-connections"]').click();
    await page.waitForSelector("#queryPane", { timeout: 5000 });
    await page.locator('[data-query-pane-action="refresh-all"]').click();
    await page.waitForFunction(
      () => document.querySelector("#statusText")?.textContent?.includes("1 件の接続を更新"),
      null,
      { timeout: 5000 },
    );
    const afterRefresh = await page.evaluate(() => ({
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector("#statusText")?.textContent?.includes("元に戻しました：すべて更新"),
      null,
      { timeout: 5000 },
    );
    const afterUndo = await page.evaluate(() => ({
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(
      () => document.querySelector("#statusText")?.textContent?.includes("やり直しました：すべて更新"),
      null,
      { timeout: 5000 },
    );
    const afterRedo = await page.evaluate(() => ({
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    return { afterRefresh, afterUndo, afterRedo };
  } finally {
    await page.close();
  }
}



async function exerciseForecastSheet(context) {
  const workbookPath = await createForecastWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?forecastSheet=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-forecast-sheet"]').click();
    await page.waitForSelector('[data-data-dialog="forecast-sheet"]', { timeout: 5000 });
    const dialogDefaults = await page.evaluate(() => ({
      forecastEnd: document.querySelector("[data-data-field='forecastEnd']")?.value || "",
      confidenceLevel: document.querySelector("[data-data-field='confidenceLevel']")?.value || "",
    }));
    await page.locator("[data-data-field='chartType'][value='column']").check();
    await page.locator("[data-data-field='forecastEnd']").fill("6");
    await page.locator("[data-data-field='confidenceLevel']").fill("80");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('.sheet-image[alt*="予測グラフ"]', { timeout: 5000 });
    const created = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      chartCount: document.querySelectorAll('.sheet-image[alt*="予測グラフ"]').length,
      headers: ["A1", "B1", "C1", "D1", "E1"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      finalPeriod: document.querySelector('.sheet-cell[data-address="A7"] .cell-content')?.textContent || "",
      confidenceLower: document.querySelector('.sheet-cell[data-address="D7"] .cell-content')?.textContent || "",
      confidenceUpper: document.querySelector('.sheet-cell[data-address="E7"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "ForecastSource",
      null,
      { timeout: 5000 },
    );
    const afterUndo = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      tabs: [...document.querySelectorAll(".sheet-tab")].map((tab) => tab.textContent.trim()),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForSelector('.sheet-image[alt*="予測グラフ"]', { timeout: 5000 });
    const afterRedo = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      chartCount: document.querySelectorAll('.sheet-image[alt*="予測グラフ"]').length,
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    return { dialogDefaults, ...created, afterUndo, afterRedo };
  } finally {
    await page.close();
  }
}



async function exerciseDataAnalysisToolpak(context) {
  const workbookPath = await createDataAnalysisWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?dataAnalysis=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="D6"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="D6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-analysis"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
    const picker = await page.evaluate(() => ({
      title: document.querySelector("#dataAnalysisPickerTitle")?.textContent || "",
      selected: document.querySelector("[data-data-analysis-picker]")?.value || "",
      methods: [...document.querySelectorAll("[data-data-analysis-picker] option")].map((option) => option.textContent.trim()),
      summary: document.querySelector("[data-data-analysis-picker-summary]")?.textContent || "",
    }));
    await page.locator("[data-data-analysis-picker]").selectOption("regression");
    await page.locator('[data-data-dialog-action="data-analysis-help"]').click();
    await page.waitForSelector("#messageBox", { timeout: 5000 });
    const pickerHelp = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      text: document.querySelector("#messageBoxText")?.textContent || "",
    }));
    await page.locator("[data-message-box-close]").click();
    await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
    const regressionDialog = await page.evaluate(() => ({
      title: document.querySelector("#dataAnalysisTitle")?.textContent || "",
      hasMethodList: Boolean(document.querySelector('[data-data-dialog="data-analysis"] .data-analysis-list')),
      visibleFields: [...document.querySelectorAll('[data-data-dialog="data-analysis"] .excel-dialog-field:not(.hidden) .excel-dialog-caption')]
        .map((item) => item.textContent.trim()),
      summary: document.querySelector("[data-data-analysis-summary]")?.textContent || "",
    }));
    await page.locator("[data-data-field='inputYRange']").fill("'分析'!B1:B6");
    await page.locator("[data-data-field='inputXRange']").fill("'分析'!A1:A6");
    await page.locator("[data-data-field='labels']").check();
    await page.locator("[data-data-field='residuals']").check();
    await page.locator("[data-data-field='standardizedResiduals']").check();
    await page.locator("[data-data-field='residualPlot']").check();
    await page.locator("[data-data-field='lineFitPlot']").check();
    await page.locator("input[name='dataAnalysisOutputMode'][value='newWorksheet']").check();
    await page.locator("[data-data-field='outputSheetName']").fill("保持確認");
    await page.locator('[data-data-dialog-action="data-analysis-help"]').click();
    await page.waitForSelector("#messageBox", { timeout: 5000 });
    const optionsHelp = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      text: document.querySelector("#messageBoxText")?.textContent || "",
    }));
    await page.locator("[data-message-box-close]").click();
    await page.locator('[data-data-dialog-action="data-analysis-options-back"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
    const pickerAfterBack = await page.evaluate(() => ({
      selected: document.querySelector("[data-data-analysis-picker]")?.value || "",
      summary: document.querySelector("[data-data-analysis-picker-summary]")?.textContent || "",
    }));
    await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
    const regressionStatePreserved = await page.evaluate(() => ({
      outputMode: document.querySelector("input[name='dataAnalysisOutputMode']:checked")?.value || "",
      outputSheetName: document.querySelector("[data-data-field='outputSheetName']")?.value || "",
      labels: Boolean(document.querySelector("[data-data-field='labels']")?.checked),
      residuals: Boolean(document.querySelector("[data-data-field='residuals']")?.checked),
      standardizedResiduals: Boolean(document.querySelector("[data-data-field='standardizedResiduals']")?.checked),
      residualPlot: Boolean(document.querySelector("[data-data-field='residualPlot']")?.checked),
      lineFitPlot: Boolean(document.querySelector("[data-data-field='lineFitPlot']")?.checked),
    }));
    await page.locator("[data-data-field='outputSheetName']").fill("回帰結果");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "回帰結果",
      null,
      { timeout: 5000 },
    );
    const regression = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      title: document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent || "",
      statLabel: document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent || "",
      coefficientHeader: document.querySelector('.sheet-cell[data-address="B15"] .cell-content')?.textContent || "",
      coefficientFixedLowerHeader: document.querySelector('.sheet-cell[data-address="H15"] .cell-content')?.textContent || "",
      coefficientFixedUpperHeader: document.querySelector('.sheet-cell[data-address="I15"] .cell-content')?.textContent || "",
      residualHeader: document.querySelector('.sheet-cell[data-address="A19"] .cell-content')?.textContent || "",
      standardizedResidualHeader: document.querySelector('.sheet-cell[data-address="D20"] .cell-content')?.textContent || "",
      chartCount: window.__webSheetDevSheetModel?.("回帰結果")?.images?.length || 0,
      chartTitles: (window.__webSheetDevSheetModel?.("回帰結果")?.images || []).map((image) => image.chart?.title || image.title || image.name || ""),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-analysis"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
    await page.locator("[data-data-analysis-picker]").selectOption("random");
    await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
    const randomDialog = await page.evaluate(() => ({
      title: document.querySelector("#dataAnalysisTitle")?.textContent || "",
      visibleFields: [...document.querySelectorAll('[data-data-dialog="data-analysis"] .excel-dialog-field:not(.hidden) .excel-dialog-caption')]
        .map((item) => item.textContent.trim()),
      summary: document.querySelector("[data-data-analysis-summary]")?.textContent || "",
    }));
    await page.locator("[data-data-field='randomRows']").fill("3");
    await page.locator("[data-data-field='randomCols']").fill("2");
    await page.locator("[data-data-field='randomDistribution']").selectOption("uniform");
    await page.locator("[data-data-field='lowerBound']").fill("10");
    await page.locator("[data-data-field='upperBound']").fill("20");
    await page.locator("[data-data-field='randomSeed']").fill("7");
    await page.locator("input[name='dataAnalysisOutputMode'][value='newWorksheet']").check();
    await page.locator("[data-data-field='outputSheetName']").fill("乱数結果");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "乱数結果",
      null,
      { timeout: 5000 },
    );
    const random = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      header: document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent || "",
      values: ["A2", "B2", "A3", "B3", "A4", "B4"].map((address) => Number(document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent)),
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator(".sheet-tab", { hasText: "分析" }).click();
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "分析", null, { timeout: 5000 });
    await page.locator('.sheet-cell[data-address="C1"]').click();
    await page.locator('.sheet-cell[data-address="D6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-analysis"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
    await page.locator("[data-data-analysis-picker]").selectOption("sampling");
    await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
    await page.locator("[data-data-field='inputRange']").fill("'分析'!C1:D6");
    await page.locator("[data-data-field='labels']").check();
    await page.locator("input[name='dataAnalysisSamplingMethod'][value='periodic']").check();
    await page.locator("[data-data-field='interval']").fill("2");
    await page.locator("[data-data-field='sampleCount']").fill("3");
    await page.locator("input[name='dataAnalysisOutputMode'][value='newWorksheet']").check();
    await page.locator("[data-data-field='outputSheetName']").fill("周期標本");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "周期標本",
      null,
      { timeout: 5000 },
    );
    const samplingPeriodic = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      values: ["A1", "A2", "B2", "C2", "A3", "B3", "C3", "A4", "B4", "C4", "A5", "B5", "C5"].map((address) =>
        document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
      ),
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator(".sheet-tab", { hasText: "分析" }).click();
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "分析", null, { timeout: 5000 });
    await page.locator('.sheet-cell[data-address="C1"]').click();
    await page.locator('.sheet-cell[data-address="D6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-analysis"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
    await page.locator("[data-data-analysis-picker]").selectOption("sampling");
    await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
    await page.locator("[data-data-field='inputRange']").fill("'分析'!C1:D6");
    await page.locator("[data-data-field='labels']").check();
    await page.locator("input[name='dataAnalysisSamplingMethod'][value='random']").check();
    await page.locator("[data-data-field='sampleCount']").fill("3");
    await page.locator("[data-data-field='randomSeed']").fill("3");
    await page.locator("input[name='dataAnalysisOutputMode'][value='newWorksheet']").check();
    await page.locator("[data-data-field='outputSheetName']").fill("ランダム標本");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "ランダム標本",
      null,
      { timeout: 5000 },
    );
    const samplingRandom = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      values: ["A1", "A2", "B2", "C2", "A3", "B3", "C3", "A4", "B4", "C4", "A5", "B5", "C5"].map((address) =>
        document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
      ),
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator(".sheet-tab", { hasText: "分析" }).click();
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "分析", null, { timeout: 5000 });
    await page.locator('.sheet-cell[data-address="C1"]').click();
    await page.locator('.sheet-cell[data-address="D6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-analysis"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
    await page.locator("[data-data-analysis-picker]").selectOption("rankPercentile");
    await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
    await page.locator("[data-data-field='inputRange']").fill("'分析'!C1:D6");
    await page.locator("[data-data-field='labels']").check();
    await page.locator("input[name='dataAnalysisOutputMode'][value='newWorksheet']").check();
    await page.locator("[data-data-field='outputSheetName']").fill("順位結果");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "順位結果",
      null,
      { timeout: 5000 },
    );
    const rankPercentile = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      values: ["A1", "A2", "B2", "C2", "D2", "A3", "B3", "C3", "D3", "A7", "B7", "C7", "D7", "A9", "B9", "C9", "D9"].map((address) =>
        document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
      ),
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator(".sheet-tab", { hasText: "分析" }).click();
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "分析", null, { timeout: 5000 });
    await page.locator('.sheet-cell[data-address="C1"]').click();
    await page.locator('.sheet-cell[data-address="D6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-analysis"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
    await page.locator("[data-data-analysis-picker]").selectOption("histogram");
    await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
    await page.locator("[data-data-field='inputRange']").fill("'分析'!C1:D6");
    await page.locator("[data-data-field='binRange']").fill("'分析'!F1:F4");
    await page.locator("[data-data-field='labels']").check();
    await page.locator("[data-data-field='chartOutput']").check();
    await page.locator("input[name='dataAnalysisOutputMode'][value='newWorksheet']").check();
    await page.locator("[data-data-field='outputSheetName']").fill("ヒストグラム結果");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "ヒストグラム結果",
      null,
      { timeout: 5000 },
    );
    const histogram = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      values: ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3", "A4", "B4", "C4", "A5", "B5", "C5"].map((address) =>
        document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
      ),
      chartCount: window.__webSheetDevSheetModel?.("ヒストグラム結果")?.images?.length || 0,
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    const runAnalysisTool = async ({ method, outputSheetName, inputRange = "", selectStart = "C1", selectEnd = "D6", labels = true, setup, addresses, chartSheet = outputSheetName }) => {
      await page.locator(".sheet-tab", { hasText: "分析" }).click();
      await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "分析", null, { timeout: 5000 });
      await page.locator(`.sheet-cell[data-address="${selectStart}"]`).click();
      await page.locator(`.sheet-cell[data-address="${selectEnd}"]`).click({ modifiers: ["Shift"] });
      await page.locator('[data-ribbon-tab="data"]').click();
      await page.locator('[data-ribbon-command="data-analysis"]').click();
      await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
      await page.locator("[data-data-analysis-picker]").selectOption(method);
      await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
      await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
      if (inputRange) await page.locator("[data-data-field='inputRange']").fill(inputRange);
      const labelsControl = page.locator("[data-data-field='labels']");
      if (await labelsControl.count()) {
        const labelsDisabled = await labelsControl.isDisabled().catch(() => true);
        if (!labelsDisabled) {
          if (labels) await labelsControl.check();
          else await labelsControl.uncheck();
        }
      }
      if (setup) await setup(page);
      const dialogCaptions = await page.evaluate(() =>
        [...document.querySelectorAll('[data-data-dialog="data-analysis"] .excel-dialog-field:not(.hidden) .excel-dialog-caption')]
          .map((item) => item.textContent.trim()),
      );
      await page.locator("input[name='dataAnalysisOutputMode'][value='newWorksheet']").check();
      await page.locator("[data-data-field='outputSheetName']").fill(outputSheetName);
      await page.locator('[data-data-dialog-action="ok"]').click();
      await page.waitForFunction(
        (sheetName) => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === sheetName,
        outputSheetName,
        { timeout: 5000 },
      );
      await page.evaluate(() => {
        const host = document.querySelector("#sheetHost");
        const horizontal = document.querySelector("#horizontalScrollbar");
        if (host) {
          host.scrollLeft = 0;
          host.scrollTop = 0;
          host.dispatchEvent(new Event("scroll"));
        }
        if (horizontal) {
          horizontal.scrollLeft = 0;
          horizontal.dispatchEvent(new Event("scroll"));
        }
      });
      await page.waitForSelector('.sheet-cell[data-address="A1"]', { timeout: 10000 });
      return page.evaluate(({ cells, chartSheetName, captions }) => ({
        activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
        header: document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent || "",
        values: cells.map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
        chartCount: window.__webSheetDevSheetModel?.(chartSheetName)?.images?.length || 0,
        dialogCaptions: captions,
        status: document.querySelector("#statusText")?.textContent || "",
      }), { cells: addresses, chartSheetName: chartSheet, captions: dialogCaptions });
    };

    const correlation = await runAnalysisTool({
      method: "correlation",
      inputRange: "'分析'!C1:D6",
      outputSheetName: "相関結果",
      addresses: ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"],
    });

    const covariance = await runAnalysisTool({
      method: "covariance",
      inputRange: "'分析'!C1:D6",
      outputSheetName: "共分散結果",
      addresses: ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"],
    });

    const correlationPairwise = await runAnalysisTool({
      method: "correlation",
      inputRange: "'分析'!U1:V6",
      selectStart: "U1",
      selectEnd: "V6",
      outputSheetName: "相関ペア",
      addresses: ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"],
    });

    const covariancePairwise = await runAnalysisTool({
      method: "covariance",
      inputRange: "'分析'!U1:V6",
      selectStart: "U1",
      selectEnd: "V6",
      outputSheetName: "共分散ペア",
      addresses: ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"],
    });

    const regressionNormalProbability = await runAnalysisTool({
      method: "regression",
      labels: true,
      outputSheetName: "正規確率回帰",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='inputYRange']").fill("'分析'!B1:B6");
        await toolPage.locator("[data-data-field='inputXRange']").fill("'分析'!A1:A6");
        await toolPage.locator("[data-data-field='normalProbabilityPlot']").check();
      },
      addresses: ["A1", "A15", "B15", "H15", "I15", "A19", "A20", "B20", "A21", "B21", "A25", "B25"],
    });

    const histogramPareto = await runAnalysisTool({
      method: "histogram",
      inputRange: "'分析'!C1:D6",
      outputSheetName: "パレートヒストグラム",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='binRange']").fill("'分析'!F1:F4");
        await toolPage.locator("[data-data-field='pareto']").check();
        await toolPage.locator("[data-data-field='chartOutput']").check();
      },
      addresses: ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3", "A4", "B4", "C4", "A5", "B5", "C5"],
    });

    const movingAverage = await runAnalysisTool({
      method: "movingAverage",
      inputRange: "'分析'!C1:C6",
      selectStart: "C1",
      selectEnd: "C6",
      outputSheetName: "移動平均結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='interval']").fill("3");
        await toolPage.locator("[data-data-field='standardErrorOutput']").check();
        await toolPage.locator("[data-data-field='chartOutput']").check();
      },
      addresses: ["A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "A3", "B3", "C3", "D3", "A4", "B4", "C4", "D4", "A5", "B5", "C5", "D5", "A6", "B6", "C6", "D6"],
    });

    const exponentialSmoothing = await runAnalysisTool({
      method: "exponentialSmoothing",
      inputRange: "'分析'!C1:C6",
      selectStart: "C1",
      selectEnd: "C6",
      outputSheetName: "指数平滑結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='alpha']").fill("0.2");
        await toolPage.locator("[data-data-field='standardErrorOutput']").check();
        await toolPage.locator("[data-data-field='chartOutput']").check();
      },
      addresses: ["A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "A3", "B3", "C3", "D3", "A4", "B4", "C4", "D4", "A5", "B5", "C5", "D5", "A6", "B6", "C6", "D6"],
    });

    const fTest = await runAnalysisTool({
      method: "fTest",
      outputSheetName: "F検定結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='variable1Range']").fill("'分析'!C1:C6");
        await toolPage.locator("[data-data-field='variable2Range']").fill("'分析'!D1:D6");
        await toolPage.locator("[data-data-field='alpha']").fill("0.05");
      },
      addresses: ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3", "A6", "B6", "C6", "A7", "B7", "C7"],
    });

    const tTestEqual = await runAnalysisTool({
      method: "tTestEqual",
      outputSheetName: "t検定結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='variable1Range']").fill("'分析'!C1:C6");
        await toolPage.locator("[data-data-field='variable2Range']").fill("'分析'!D1:D6");
        await toolPage.locator("[data-data-field='hypothesizedMeanDiff']").fill("0");
        await toolPage.locator("[data-data-field='alpha']").fill("0.05");
      },
      addresses: ["A1", "B1", "C1", "A2", "B2", "C2", "A5", "B5", "C5", "A7", "B7", "C7", "A9", "B9", "C9", "A12", "B12"],
    });

    const zTest = await runAnalysisTool({
      method: "zTest",
      outputSheetName: "z検定結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='variable1Range']").fill("'分析'!C1:C6");
        await toolPage.locator("[data-data-field='variable2Range']").fill("'分析'!D1:D6");
        await toolPage.locator("[data-data-field='knownVariance1']").fill("13.6");
        await toolPage.locator("[data-data-field='knownVariance2']").fill("10.96");
        await toolPage.locator("[data-data-field='hypothesizedMeanDiff']").fill("0");
        await toolPage.locator("[data-data-field='alpha']").fill("0.05");
      },
      addresses: ["A1", "B1", "C1", "A2", "B2", "C2", "A5", "B5", "C5", "A7", "B7", "C7", "A9", "B9", "C9", "A11", "B11"],
    });

    const tTestPaired = await runAnalysisTool({
      method: "tTestPaired",
      outputSheetName: "対応t検定結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='variable1Range']").fill("'分析'!C1:C6");
        await toolPage.locator("[data-data-field='variable2Range']").fill("'分析'!D1:D6");
        await toolPage.locator("[data-data-field='hypothesizedMeanDiff']").fill("0");
        await toolPage.locator("[data-data-field='alpha']").fill("0.05");
      },
      addresses: ["A1", "B1", "C1", "A2", "B2", "C2", "A5", "B5", "C5", "A8", "B8", "C8", "A10", "B10", "C10", "A13", "B13"],
    });

    const tTestUnequal = await runAnalysisTool({
      method: "tTestUnequal",
      outputSheetName: "不等分散t検定結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='variable1Range']").fill("'分析'!C1:C6");
        await toolPage.locator("[data-data-field='variable2Range']").fill("'分析'!D1:D6");
        await toolPage.locator("[data-data-field='hypothesizedMeanDiff']").fill("0");
        await toolPage.locator("[data-data-field='alpha']").fill("0.05");
      },
      addresses: ["A1", "B1", "C1", "A2", "B2", "C2", "A5", "B5", "C5", "A7", "B7", "C7", "A9", "B9", "C9", "A12", "B12"],
    });

    const randomPattern = await runAnalysisTool({
      method: "random",
      labels: false,
      outputSheetName: "乱数パターン",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='randomRows']").fill("2");
        await toolPage.locator("[data-data-field='randomCols']").fill("3");
        await toolPage.locator("[data-data-field='randomDistribution']").selectOption("pattern");
        await toolPage.locator("[data-data-field='lowerBound']").fill("1");
        await toolPage.locator("[data-data-field='upperBound']").fill("5");
        await toolPage.locator("[data-data-field='patternStep']").fill("2");
      },
      addresses: ["A1", "A2", "B2", "C2", "A3", "B3", "C3"],
    });

    const randomDiscrete = await runAnalysisTool({
      method: "random",
      labels: false,
      outputSheetName: "乱数離散",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='randomRows']").fill("2");
        await toolPage.locator("[data-data-field='randomCols']").fill("3");
        await toolPage.locator("[data-data-field='randomDistribution']").selectOption("discrete");
        await toolPage.locator("[data-data-field='probabilityRange']").fill("'分析'!A7:B10");
        await toolPage.locator("[data-data-field='randomSeed']").fill("5");
      },
      addresses: ["A1", "A2", "B2", "C2", "A3", "B3", "C3"],
    });

    const randomBinomial = await runAnalysisTool({
      method: "random",
      labels: false,
      outputSheetName: "乱数二項",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='randomRows']").fill("2");
        await toolPage.locator("[data-data-field='randomCols']").fill("3");
        await toolPage.locator("[data-data-field='randomDistribution']").selectOption("binomial");
        await toolPage.locator("[data-data-field='mean']").fill("0.5");
        await toolPage.locator("[data-data-field='stddev']").fill("4");
        await toolPage.locator("[data-data-field='randomSeed']").fill("2");
      },
      addresses: ["A1", "A2", "B2", "C2", "A3", "B3", "C3"],
    });

    const randomPoisson = await runAnalysisTool({
      method: "random",
      labels: false,
      outputSheetName: "乱数ポアソン",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='randomRows']").fill("2");
        await toolPage.locator("[data-data-field='randomCols']").fill("3");
        await toolPage.locator("[data-data-field='randomDistribution']").selectOption("poisson");
        await toolPage.locator("[data-data-field='mean']").fill("3");
        await toolPage.locator("[data-data-field='randomSeed']").fill("11");
      },
      addresses: ["A1", "A2", "B2", "C2", "A3", "B3", "C3"],
    });

    const randomNormal = await runAnalysisTool({
      method: "random",
      labels: false,
      outputSheetName: "乱数正規",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='randomRows']").fill("2");
        await toolPage.locator("[data-data-field='randomCols']").fill("3");
        await toolPage.locator("[data-data-field='randomDistribution']").selectOption("normal");
        await toolPage.locator("[data-data-field='mean']").fill("10");
        await toolPage.locator("[data-data-field='stddev']").fill("2");
        await toolPage.locator("[data-data-field='randomSeed']").fill("17");
      },
      addresses: ["A1", "A2", "B2", "C2", "A3", "B3", "C3"],
    });

    const randomBernoulli = await runAnalysisTool({
      method: "random",
      labels: false,
      outputSheetName: "乱数ベルヌーイ",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='randomRows']").fill("2");
        await toolPage.locator("[data-data-field='randomCols']").fill("3");
        await toolPage.locator("[data-data-field='randomDistribution']").selectOption("bernoulli");
        await toolPage.locator("[data-data-field='mean']").fill("0.5");
        await toolPage.locator("[data-data-field='randomSeed']").fill("13");
      },
      addresses: ["A1", "A2", "B2", "C2", "A3", "B3", "C3"],
    });

    const descriptive = await runAnalysisTool({
      method: "descriptive",
      inputRange: "'分析'!C1:D6",
      outputSheetName: "記述統計結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='confidence']").fill("95");
        await toolPage.locator("[data-data-field='kthLargestEnabled']").check();
        await toolPage.locator("[data-data-field='kthLargest']").fill("2");
        await toolPage.locator("[data-data-field='kthSmallestEnabled']").check();
        await toolPage.locator("[data-data-field='kthSmallest']").fill("2");
      },
      addresses: [
        "A1", "B1", "C1",
        "A2", "B2", "C2",
        "A3", "B3", "C3",
        "A7", "B7", "C7",
        "A14", "B14", "C14",
        "A15", "B15", "C15",
        "A16", "B16", "C16",
        "A17", "B17", "C17",
      ],
    });

    const anovaOneWay = await runAnalysisTool({
      method: "anova1",
      inputRange: "'分析'!C1:D6",
      outputSheetName: "一元分散結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='alpha']").fill("0.05");
      },
      addresses: [
        "A1", "B1", "C1", "D1", "E1",
        "A2", "B2", "C2", "D2", "E2",
        "A3", "B3", "C3", "D3", "E3",
        "A5",
        "A6", "B6", "C6", "D6", "E6", "F6", "G6",
        "A7", "B7", "C7", "D7", "E7", "F7", "G7",
      ],
    });

    const anovaTwoFactor = await runAnalysisTool({
      method: "anova2NoReplication",
      inputRange: "'分析'!M1:P4",
      selectStart: "M1",
      selectEnd: "P4",
      outputSheetName: "二元分散結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='alpha']").fill("0.05");
      },
      addresses: [
        "A1", "B1", "C1", "D1", "E1",
        "A2", "B2", "C2", "D2", "E2",
        "A3", "B3", "C3", "D3", "E3",
        "A6", "B6", "C6", "D6", "E6",
        "A9",
        "A10", "B10", "C10", "D10", "E10", "F10", "G10",
        "A11", "B11", "C11", "D11", "E11", "F11", "G11",
      ],
    });

    const anovaReplication = await runAnalysisTool({
      method: "anova2Replication",
      inputRange: "'分析'!R1:T5",
      selectStart: "R1",
      selectEnd: "T5",
      outputSheetName: "繰返分散結果",
      setup: async (toolPage) => {
        await toolPage.locator("[data-data-field='interval']").fill("2");
        await toolPage.locator("[data-data-field='alpha']").fill("0.05");
      },
      addresses: [
        "A1", "B1", "C1", "D1",
        "A2",
        "A3", "B3", "C3", "D3",
        "A4", "B4", "C4", "D4",
        "A7",
        "A8", "B8", "C8", "D8",
        "A12",
        "A13", "B13", "C13", "D13", "E13", "F13", "G13",
        "A14", "B14", "C14", "D14", "E14", "F14", "G14",
      ],
    });

    await page.locator(".sheet-tab", { hasText: "分析" }).click();
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "分析", null, { timeout: 5000 });
    await page.locator('.sheet-cell[data-address="H1"]').click();
    await page.locator('.sheet-cell[data-address="H5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-analysis"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
    await page.locator("[data-data-analysis-picker]").selectOption("fourier");
    await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
    await page.locator("[data-data-field='inputRange']").fill("'分析'!H1:H5");
    await page.locator("[data-data-field='labels']").check();
    await page.locator("input[name='dataAnalysisOutputMode'][value='newWorksheet']").check();
    await page.locator("[data-data-field='outputSheetName']").fill("フーリエ結果");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "フーリエ結果",
      null,
      { timeout: 5000 },
    );
    const fourierReal = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      values: ["A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "A3", "B3", "C3", "D3", "A4", "B4", "C4", "D4", "A5", "B5", "C5", "D5"].map((address) =>
        document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
      ),
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator(".sheet-tab", { hasText: "分析" }).click();
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "分析", null, { timeout: 5000 });
    await page.locator('.sheet-cell[data-address="I1"]').click();
    await page.locator('.sheet-cell[data-address="I5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-analysis"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
    await page.locator("[data-data-analysis-picker]").selectOption("fourier");
    await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
    await page.locator("[data-data-field='inputRange']").fill("'分析'!I1:I5");
    await page.locator("[data-data-field='labels']").check();
    await page.locator("input[name='dataAnalysisOutputMode'][value='newWorksheet']").check();
    await page.locator("[data-data-field='outputSheetName']").fill("複素フーリエ");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "複素フーリエ",
      null,
      { timeout: 5000 },
    );
    const fourierComplex = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      values: ["A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "A3", "B3", "C3", "D3", "A4", "B4", "C4", "D4", "A5", "B5", "C5", "D5"].map((address) =>
        document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
      ),
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator(".sheet-tab", { hasText: "分析" }).click();
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "分析", null, { timeout: 5000 });
    await page.locator('.sheet-cell[data-address="K1"]').click();
    await page.locator('.sheet-cell[data-address="K5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-analysis"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis-picker"]', { timeout: 5000 });
    await page.locator("[data-data-analysis-picker]").selectOption("fourier");
    await page.locator('[data-data-dialog-action="data-analysis-picker-ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-analysis"]', { timeout: 5000 });
    await page.locator("[data-data-field='inputRange']").fill("'分析'!K1:K5");
    await page.locator("[data-data-field='labels']").check();
    await page.locator("[data-data-field='inverseFourier']").check();
    await page.locator("input[name='dataAnalysisOutputMode'][value='newWorksheet']").check();
    await page.locator("[data-data-field='outputSheetName']").fill("逆フーリエ");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "逆フーリエ",
      null,
      { timeout: 5000 },
    );
    const fourierInverse = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      values: ["A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "A3", "B3", "C3", "D3", "A4", "B4", "C4", "D4", "A5", "B5", "C5", "D5"].map((address) =>
        document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
      ),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    return { picker, pickerHelp, optionsHelp, pickerAfterBack, regressionDialog, regressionStatePreserved, regression, randomDialog, random, samplingPeriodic, samplingRandom, rankPercentile, histogram, correlation, covariance, correlationPairwise, covariancePairwise, regressionNormalProbability, histogramPareto, movingAverage, exponentialSmoothing, fTest, tTestEqual, zTest, tTestPaired, tTestUnequal, randomPattern, randomDiscrete, randomBinomial, randomPoisson, randomNormal, randomBernoulli, descriptive, anovaOneWay, anovaTwoFactor, anovaReplication, fourierReal, fourierComplex, fourierInverse };
  } finally {
    await page.close();
  }
}



async function exerciseConsolidateDialog(context) {
  const workbookPath = await createConsolidateWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?consolidate=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="G3"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="C3"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-consolidate"]').click();
    await page.waitForSelector('[data-data-dialog="consolidate"]', { timeout: 5000 });
    await page.locator("[data-data-field='reference']").fill("E1:G3");
    await page.locator('[data-data-dialog-action="add-reference"]').click();
    await page.locator("[data-data-field='leftColumn']").check();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('.sheet-cell[data-address="C8"]', { timeout: 5000 });
    const staticResult = await page.evaluate(() => ({
      headers: ["A5", "B5", "C5"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      rows: [6, 7, 8].map((row) => ({
        label: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        q1: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
        q2: document.querySelector(`.sheet-cell[data-address="C${row}"] .cell-content`)?.textContent || "",
      })),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.close();

    const linkPage = await context.newPage();
    try {
      await linkPage.goto(`${appUrl}?consolidateLinks=${Date.now()}`, { waitUntil: "load" });
      await linkPage.setInputFiles("#fileInput", workbookPath);
      await waitForUploadedWorkbook(linkPage, workbookPath);
      await linkPage.waitForSelector('.sheet-cell[data-address="C3"]', { timeout: 10000 });
      await linkPage.locator('.sheet-cell[data-address="A1"]').click();
      await linkPage.locator('.sheet-cell[data-address="C3"]').click({ modifiers: ["Shift"] });
      await linkPage.locator('[data-ribbon-tab="data"]').click();
      await linkPage.locator('[data-ribbon-command="data-consolidate"]').click();
      await linkPage.waitForSelector('[data-data-dialog="consolidate"]', { timeout: 5000 });
      const initialReference = await linkPage.locator("[data-data-field='references'] option").first().textContent();
      const functionOptions = await linkPage.locator("[data-data-field='function'] option").evaluateAll((options) => options.map((option) => option.textContent.trim()));
      await linkPage.locator("[data-data-field='reference']").fill("ConsolidateExtra");
      await linkPage.locator('[data-data-dialog-action="add-reference"]').click();
      const references = await linkPage.evaluate(() =>
        [...document.querySelectorAll("[data-data-field='references'] option")].map((option) => option.textContent || ""),
      );
      await linkPage.locator("[data-data-field='leftColumn']").check();
      await linkPage.locator("[data-data-field='createLinks']").check();
      await linkPage.locator('[data-data-dialog-action="ok"]').click();
      await linkPage.waitForFunction(
        () => document.querySelector('.sheet-cell[data-address="B6"] .cell-content')?.textContent === "21",
        null,
        { timeout: 5000 },
      );
      await linkPage.locator('.sheet-cell[data-address="B6"]').click();
      const linked = await linkPage.evaluate(() => ({
        headers: ["A5", "B5", "C5"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
        rows: [6, 7, 8].map((row) => ({
          label: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
          q1: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
          q2: document.querySelector(`.sheet-cell[data-address="C${row}"] .cell-content`)?.textContent || "",
        })),
        formula: document.querySelector("#formulaInput")?.value || "",
        status: document.querySelector("#statusText")?.textContent || "",
      }));
      linked.initialReference = initialReference;
      linked.references = references;
      linked.functionOptions = functionOptions;
      return { static: staticResult, linked };
    } finally {
      await linkPage.close();
    }
  } finally {
    if (!page.isClosed()) await page.close();
  }
}



async function exerciseSubtotalDialog(context) {
  const workbookPath = await createSubtotalWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?subtotal=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="B3"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-subtotal"]').click();
    await page.waitForSelector('[data-data-dialog="subtotal"]', { timeout: 5000 });
    const defaults = await page.evaluate(() => ({
      groupBy: document.querySelector("[data-data-field='groupBy'] option:checked")?.textContent || "",
      functionLabel: document.querySelector("[data-data-field='function'] option:checked")?.textContent || "",
      fields: [...document.querySelectorAll("[data-data-subtotal-col]:checked")].map((input) => input.closest("label")?.textContent?.trim() || ""),
    }));
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A8"] .cell-content')?.textContent === "総計",
      null,
      { timeout: 5000 },
    );
    const rows = await page.evaluate(() =>
      [1, 2, 3, 4, 5, 6, 7, 8].map((row) => ({
        row,
        group: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        amount: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const status = await page.locator("#statusText").innerText();
    await page.locator('.sheet-cell[data-address="B3"]').click();
    await page.locator('[data-ribbon-command="data-subtotal"]').click();
    await page.waitForSelector('[data-data-dialog="subtotal"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="remove-all"]').click();
    await page.waitForFunction(
      () =>
        document.querySelector('.sheet-cell[data-address="A4"] .cell-content')?.textContent === "B" &&
        document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent !== "B 合計",
      null,
      { timeout: 5000 },
    );
    const removeAll = await page.evaluate(() => ({
      rows: [1, 2, 3, 4, 5].map((row) => ({
        row,
        group: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        amount: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator('.sheet-cell[data-address="B3"]').click();
    await page.locator('[data-ribbon-command="data-subtotal"]').click();
    await page.waitForSelector('[data-data-dialog="subtotal"]', { timeout: 5000 });
    await page.locator("[data-data-field='summaryBelow']").uncheck();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "総計",
      null,
      { timeout: 5000 },
    );
    const summaryAbove = await page.evaluate(() => ({
      rows: [1, 2, 3, 4, 5, 6, 7, 8].map((row) => ({
        row,
        group: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        amount: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    return { defaults, rows, status, removeAll, summaryAbove };
  } finally {
    await page.close();
  }
}



async function exerciseOutlineGrouping(context) {
  const workbookPath = await createOutlineWorkbook();
  const page = await context.newPage();
  const outlineRibbonState = () => page.evaluate(() => {
    const stateFor = (command) => {
      const button = document.querySelector(`[data-ribbon-command="${command}"]`);
      return {
        active: button?.classList.contains("is-active") || false,
        disabled: button?.disabled || false,
      };
    };
    return {
      show: stateFor("data-show-detail"),
      hide: stateFor("data-hide-detail"),
    };
  });
  const rowHidden = (row) => page.evaluate((targetRow) => {
    const header = document.querySelector(`.row-header[data-row="${targetRow}"]`);
    return !header || getComputedStyle(header).display === "none";
  }, row);
  const colHidden = (col) => page.evaluate((targetCol) => {
    const header = document.querySelector(`.column-header[data-col="${targetCol}"]`);
    return !header || getComputedStyle(header).display === "none" || header.getBoundingClientRect().width <= 10;
  }, col);
  try {
    await page.goto(`${appUrl}?outline=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    const importedOutlineModel = await page.evaluate(() => {
      const sheet = window.__webSheetDevSheetModel?.("ImportedOutline") || {};
      const row = (index) => sheet.rows?.[index - 1] || {};
      const col = (index) => sheet.columns?.[index - 1] || {};
      return {
        rowCount: sheet.rowCount || 0,
        colCount: sheet.colCount || 0,
        row9Hidden: row(9).hidden === true,
        row9OutlineLevel: row(9).outlineLevel || 0,
        row10Collapsed: row(10).outlineCollapsed === true,
        col8Hidden: col(8).hidden === true,
        col8OutlineLevel: col(8).outlineLevel || 0,
        col9Hidden: col(9).hidden === true,
        col10Collapsed: col(10).outlineCollapsed === true,
        applyStyles: sheet.outlineSettings?.applyStyles === true,
      };
    });
    await page.evaluate(() => window.__webSheetDevSetActiveSheet?.("ImportedOutline"));
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "ImportedOutline", null, { timeout: 5000 });
    await page.waitForSelector('[data-outline-toggle="rows"][data-outline-start="9"][data-outline-end="9"][data-outline-collapsed="true"]', { timeout: 5000 });
    await page.waitForSelector('[data-outline-toggle="columns"][data-outline-start="8"][data-outline-end="9"][data-outline-collapsed="true"]', { timeout: 5000 });
    const importedOutlineUi = {
      rowToggleText: await page.locator('[data-outline-toggle="rows"][data-outline-start="9"][data-outline-end="9"][data-outline-collapsed="true"]').first().innerText(),
      columnToggleText: await page.locator('[data-outline-toggle="columns"][data-outline-start="8"][data-outline-end="9"][data-outline-collapsed="true"]').first().innerText(),
      rowLevelButtons: await page.locator('[data-outline-level-command="rows"]').evaluateAll((buttons) => buttons.map((button) => button.textContent.trim())),
      columnLevelButtons: await page.locator('[data-outline-level-command="columns"]').evaluateAll((buttons) => buttons.map((button) => button.textContent.trim())),
      activeRowLevel: await page.locator('[data-outline-level-command="rows"].is-active').first().innerText(),
      activeColumnLevel: await page.locator('[data-outline-level-command="columns"].is-active').first().innerText(),
      row9Hidden: await rowHidden(9),
      row10Hidden: await rowHidden(10),
      col8Hidden: await colHidden(8),
      col9Hidden: await colHidden(9),
      col10Hidden: await colHidden(10),
    };
    await page.locator('[data-outline-level-command="rows"][data-outline-display-level="2"]').click();
    await page.waitForFunction(() => {
      const header = document.querySelector('.row-header[data-row="9"]');
      return header && getComputedStyle(header).display !== "none";
    }, null, { timeout: 5000 });
    await page.locator('[data-outline-level-command="columns"][data-outline-display-level="2"]').click();
    await page.waitForFunction(() => [8, 9].every((col) => {
      const header = document.querySelector(`.column-header[data-col="${col}"]`);
      return header && getComputedStyle(header).display !== "none" && header.getBoundingClientRect().width > 10;
    }), null, { timeout: 5000 });
    const importedOutlineLevelExpanded = {
      activeRowLevel: await page.locator('[data-outline-level-command="rows"].is-active').first().innerText(),
      activeColumnLevel: await page.locator('[data-outline-level-command="columns"].is-active').first().innerText(),
      row9Hidden: await rowHidden(9),
      col8Hidden: await colHidden(8),
      col9Hidden: await colHidden(9),
    };
    await page.locator('[data-outline-level-command="rows"][data-outline-display-level="1"]').click();
    await page.waitForSelector('[data-outline-toggle="rows"][data-outline-start="9"][data-outline-end="9"][data-outline-collapsed="true"]', { timeout: 5000 });
    await page.locator('[data-outline-level-command="columns"][data-outline-display-level="1"]').click();
    await page.waitForSelector('[data-outline-toggle="columns"][data-outline-start="8"][data-outline-end="9"][data-outline-collapsed="true"]', { timeout: 5000 });
    const importedOutlineLevelCollapsed = {
      activeRowLevel: await page.locator('[data-outline-level-command="rows"].is-active').first().innerText(),
      activeColumnLevel: await page.locator('[data-outline-level-command="columns"].is-active').first().innerText(),
      row9Hidden: await rowHidden(9),
      col8Hidden: await colHidden(8),
      col9Hidden: await colHidden(9),
    };
    const importedOutlineExport = await exportedOutlineWorkbookState(
      await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
      "ImportedOutline",
    );
    await page.evaluate(() => window.__webSheetDevSetActiveSheet?.("Outline"));
    await page.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "Outline", null, { timeout: 5000 });
    await page.waitForSelector('.sheet-cell[data-address="E6"]', { timeout: 10000 });
    await page.locator('.row-header[data-row="2"]').click();
    await page.locator('.row-header[data-row="4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-group"]').click();
    await page.waitForSelector("#dataCommandMenu.is-outline-group", { timeout: 5000 });
    const groupMenuActions = await page.locator("#dataCommandMenu [data-data-menu-action]").evaluateAll((buttons) =>
      buttons.map((button) => button.getAttribute("data-data-menu-action")),
    );
    await page.locator('[data-data-menu-action="outline-group"]').click();
    await page.waitForSelector('[data-data-dialog="outline-group"]', { timeout: 5000 });
    const rowGroupDefault = await page.locator("[data-data-field='axis']:checked").inputValue();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-outline-toggle="rows"][data-outline-start="2"][data-outline-end="4"]', { timeout: 5000 });
    const outlineRibbonAfterGroup = await outlineRibbonState();
    const rowToggleBefore = await page.locator('[data-outline-toggle="rows"][data-outline-start="2"][data-outline-end="4"]').first().innerText();
    await page.locator('[data-outline-toggle="rows"][data-outline-start="2"][data-outline-end="4"]').first().click();
    await page.waitForSelector('[data-outline-toggle="rows"][data-outline-start="2"][data-outline-end="4"][data-outline-collapsed="true"]', { timeout: 5000 });
    const rowsCollapsed = {
      row2: await rowHidden(2),
      row3: await rowHidden(3),
      row4: await rowHidden(4),
      row5: await rowHidden(5),
      toggleText: await page.locator('[data-outline-toggle="rows"][data-outline-start="2"][data-outline-end="4"][data-outline-collapsed="true"]').first().innerText(),
    };
    const rowCollapsedExport = await exportedOutlineWorkbookState(
      await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
      "Outline",
    );
    await page.locator('[data-outline-toggle="rows"][data-outline-start="2"][data-outline-end="4"][data-outline-collapsed="true"]').first().click();
    await page.waitForFunction(
      () => [...document.querySelectorAll(".row-header[data-row]")].slice(0, 6).every((node) => getComputedStyle(node).display !== "none"),
      null,
      { timeout: 5000 },
    );
    const rowsExpanded = {
      row2: await rowHidden(2),
      row3: await rowHidden(3),
      row4: await rowHidden(4),
    };
    await page.locator('.row-header[data-row="2"]').click();
    await page.locator('.row-header[data-row="4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-ungroup"]').click();
    await page.waitForSelector("#dataCommandMenu.is-outline-ungroup", { timeout: 5000 });
    const ungroupMenuActions = await page.locator("#dataCommandMenu [data-data-menu-action]").evaluateAll((buttons) =>
      buttons.map((button) => button.getAttribute("data-data-menu-action")),
    );
    await page.locator('[data-data-menu-action="outline-ungroup"]').click();
    await page.waitForSelector('[data-data-dialog="outline-ungroup"]', { timeout: 5000 });
    const rowUngroupDefault = await page.locator("[data-data-field='axis']:checked").inputValue();
    const rowUngroupNote = await page.locator(".outline-dialog .data-dialog-note").innerText();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(() => !document.querySelector('[data-outline-toggle="rows"]'), null, { timeout: 5000 });
    const rowTogglesAfterUngroup = await page.locator('[data-outline-toggle="rows"]').count();

    await page.locator('.column-header[data-col="2"]').click();
    await page.locator('.column-header[data-col="4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-group"]').click();
    await page.waitForSelector("#dataCommandMenu.is-outline-group", { timeout: 5000 });
    await page.locator('[data-data-menu-action="outline-group"]').click();
    await page.waitForSelector('[data-data-dialog="outline-group"]', { timeout: 5000 });
    const columnGroupDefault = await page.locator("[data-data-field='axis']:checked").inputValue();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-outline-toggle="columns"][data-outline-start="2"][data-outline-end="4"]', { timeout: 5000 });
    const columnToggleBefore = await page.locator('[data-outline-toggle="columns"][data-outline-start="2"][data-outline-end="4"]').first().innerText();
    await page.locator('[data-outline-toggle="columns"][data-outline-start="2"][data-outline-end="4"]').first().click();
    await page.waitForSelector('[data-outline-toggle="columns"][data-outline-start="2"][data-outline-end="4"][data-outline-collapsed="true"]', { timeout: 5000 });
    await page.waitForFunction(
      () => [2, 3, 4].every((col) => {
        const header = document.querySelector(`.column-header[data-col="${col}"]`);
        return !header || getComputedStyle(header).display === "none" || header.getBoundingClientRect().width <= 10;
      }),
      null,
      { timeout: 5000 },
    );
    const columnsCollapsed = {
      col2: await colHidden(2),
      col3: await colHidden(3),
      col4: await colHidden(4),
      col5: await colHidden(5),
      toggleText: await page.locator('[data-outline-toggle="columns"][data-outline-start="2"][data-outline-end="4"][data-outline-collapsed="true"]').first().innerText(),
    };
    const columnCollapsedExport = await exportedOutlineWorkbookState(
      await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
      "Outline",
    );
    await page.locator('[data-outline-toggle="columns"][data-outline-start="2"][data-outline-end="4"][data-outline-collapsed="true"]').first().click();
    await page.waitForFunction(
      () => [2, 3, 4].every((col) => {
        const header = document.querySelector(`.column-header[data-col="${col}"]`);
        return header && getComputedStyle(header).display !== "none" && header.getBoundingClientRect().width > 10;
      }),
      null,
      { timeout: 5000 },
    );
    const columnsExpanded = {
      col2: await colHidden(2),
      col3: await colHidden(3),
      col4: await colHidden(4),
    };
    await page.locator('[data-ribbon-command="data-group"]').click();
    await page.waitForSelector("#dataCommandMenu.is-outline-group", { timeout: 5000 });
    await page.locator('[data-data-menu-action="outline-settings"]').click();
    await page.waitForSelector('[data-data-dialog="outline-settings"]', { timeout: 5000 });
    const outlineSettingsDefaults = await page.evaluate(() => ({
      summaryBelow: document.querySelector("[data-data-field='summaryBelow']")?.checked,
      summaryRight: document.querySelector("[data-data-field='summaryRight']")?.checked,
      autoStyle: document.querySelector("[data-data-field='autoStyle']")?.checked,
    }));
    await page.locator("[data-data-field='summaryBelow']").uncheck();
    await page.locator("[data-data-field='summaryRight']").uncheck();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.locator('.row-header[data-row="2"]').click();
    await page.locator('.row-header[data-row="4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-group"]').click();
    await page.waitForSelector("#dataCommandMenu.is-outline-group", { timeout: 5000 });
    await page.locator('[data-data-menu-action="outline-group"]').click();
    await page.waitForSelector('[data-data-dialog="outline-group"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-outline-toggle="rows"][data-outline-start="2"][data-outline-end="4"]', { timeout: 5000 });
    const summaryAboveRowToggleHost = await page.locator('[data-outline-toggle="rows"][data-outline-start="2"][data-outline-end="4"]').first().evaluate((node) =>
      node.closest(".row-header")?.getAttribute("data-row"),
    );
    await page.locator('[data-ribbon-command="data-ungroup"]').click();
    await page.waitForSelector("#dataCommandMenu.is-outline-ungroup", { timeout: 5000 });
    await page.locator('[data-data-menu-action="outline-clear"]').click();
    await page.waitForFunction(() => !document.querySelector('[data-outline-toggle]'), null, { timeout: 5000 });
    const outlineCleared = await page.locator("[data-outline-toggle]").count();
    const outlineRibbonAfterClear = await outlineRibbonState();
    await page.locator('[data-ribbon-command="data-group"]').click();
    await page.waitForSelector("#dataCommandMenu.is-outline-group", { timeout: 5000 });
    await page.locator('[data-data-menu-action="outline-settings"]').click();
    await page.waitForSelector('[data-data-dialog="outline-settings"]', { timeout: 5000 });
    await page.locator("[data-data-field='summaryBelow']").check();
    await page.locator("[data-data-field='summaryRight']").check();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.locator('.sheet-cell[data-address="A2"]').click();
    await page.locator('.sheet-cell[data-address="E6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-group"]').click();
    await page.waitForSelector("#dataCommandMenu.is-outline-group", { timeout: 5000 });
    await page.locator('[data-data-menu-action="outline-auto"]').click();
    await page.waitForSelector('[data-outline-toggle="rows"]', { timeout: 5000 });
    const autoOutlineRows = await page.locator('[data-outline-toggle="rows"]').count();
    const status = await page.locator("#statusText").innerText();
    const exportFailures = [];
    if (importedOutlineModel.rowCount < 10) exportFailures.push(`import row extent ${importedOutlineModel.rowCount}`);
    if (importedOutlineModel.colCount < 10) exportFailures.push(`import column extent ${importedOutlineModel.colCount}`);
    if (importedOutlineModel.applyStyles !== true) exportFailures.push(`import outline settings ${JSON.stringify(importedOutlineModel)}`);
    if (!importedOutlineModel.row9Hidden || importedOutlineModel.row9OutlineLevel !== 1 || !importedOutlineModel.row10Collapsed) exportFailures.push(`import row outline ${JSON.stringify(importedOutlineModel)}`);
    if (!importedOutlineModel.col8Hidden || importedOutlineModel.col8OutlineLevel !== 1 || !importedOutlineModel.col9Hidden || !importedOutlineModel.col10Collapsed) exportFailures.push(`import column outline ${JSON.stringify(importedOutlineModel)}`);
    if (importedOutlineUi.rowToggleText !== "+" || importedOutlineUi.columnToggleText !== "+") exportFailures.push(`import outline toggles ${JSON.stringify(importedOutlineUi)}`);
    if (importedOutlineUi.rowLevelButtons.join("|") !== "1|2" || importedOutlineUi.columnLevelButtons.join("|") !== "1|2") exportFailures.push(`import outline level buttons ${JSON.stringify(importedOutlineUi)}`);
    if (importedOutlineUi.activeRowLevel !== "1" || importedOutlineUi.activeColumnLevel !== "1") exportFailures.push(`import active outline levels ${JSON.stringify(importedOutlineUi)}`);
    if (!importedOutlineUi.row9Hidden || importedOutlineUi.row10Hidden || !importedOutlineUi.col8Hidden || !importedOutlineUi.col9Hidden || importedOutlineUi.col10Hidden) exportFailures.push(`import hidden UI ${JSON.stringify(importedOutlineUi)}`);
    if (importedOutlineLevelExpanded.activeRowLevel !== "2" || importedOutlineLevelExpanded.activeColumnLevel !== "2" || importedOutlineLevelExpanded.row9Hidden || importedOutlineLevelExpanded.col8Hidden || importedOutlineLevelExpanded.col9Hidden) exportFailures.push(`import level expanded ${JSON.stringify(importedOutlineLevelExpanded)}`);
    if (importedOutlineLevelCollapsed.activeRowLevel !== "1" || importedOutlineLevelCollapsed.activeColumnLevel !== "1" || !importedOutlineLevelCollapsed.row9Hidden || !importedOutlineLevelCollapsed.col8Hidden || !importedOutlineLevelCollapsed.col9Hidden) exportFailures.push(`import level collapsed ${JSON.stringify(importedOutlineLevelCollapsed)}`);
    if (importedOutlineExport.rowHidden.join("|") !== "9" || importedOutlineExport.rowCollapsed.join("|") !== "10") exportFailures.push(`import row export ${JSON.stringify(importedOutlineExport)}`);
    if (importedOutlineExport.columnHidden.join("|") !== "8|9" || importedOutlineExport.columnCollapsed.join("|") !== "10") exportFailures.push(`import column export ${JSON.stringify(importedOutlineExport)}`);
    if (importedOutlineExport.outlinePr.applyStyles !== "1") exportFailures.push(`import outline settings export ${JSON.stringify(importedOutlineExport)}`);
    if (rowCollapsedExport.rowHidden.join("|") !== "2|3|4") exportFailures.push(`row hidden export ${rowCollapsedExport.rowHidden.join("|")}`);
    if (rowCollapsedExport.rowCollapsed.join("|") !== "5") exportFailures.push(`row collapsed export ${rowCollapsedExport.rowCollapsed.join("|")}`);
    if (columnCollapsedExport.columnHidden.join("|") !== "2|3|4") exportFailures.push(`column hidden export ${columnCollapsedExport.columnHidden.join("|")}`);
    if (columnCollapsedExport.columnCollapsed.join("|") !== "5") exportFailures.push(`column collapsed export ${columnCollapsedExport.columnCollapsed.join("|")}`);
    if (exportFailures.length) throw new Error(`Outline collapsed export failed: ${exportFailures.join(", ")}`);
    return {
      groupMenuActions,
      ungroupMenuActions,
      importedOutlineModel,
      importedOutlineUi,
      importedOutlineLevelExpanded,
      importedOutlineLevelCollapsed,
      importedOutlineExport,
      outlineRibbonAfterGroup,
      rowGroupDefault,
      rowToggleBefore,
      rowsCollapsed,
      rowCollapsedExport,
      rowsExpanded,
      rowUngroupDefault,
      rowUngroupNote,
      rowTogglesAfterUngroup,
      columnGroupDefault,
      columnToggleBefore,
      columnsCollapsed,
      columnCollapsedExport,
      columnsExpanded,
      outlineSettingsDefaults,
      summaryAboveRowToggleHost,
      outlineCleared,
      outlineRibbonAfterClear,
      autoOutlineRows,
      status,
    };
  } finally {
    await page.close();
  }
}



async function exerciseTextToColumnsDialog(context) {
  const workbookPath = await createTextToColumnsWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?textColumns=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="A3"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="A3"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-text-to-columns"]').click();
    await page.waitForSelector('[data-data-dialog="text-to-columns"]', { timeout: 5000 });
    const delimiterDefaults = await page.evaluate(() => ({
      tab: document.querySelector("[data-data-field='tab']")?.checked,
      semicolon: document.querySelector("[data-data-field='semicolon']")?.checked,
      comma: document.querySelector("[data-data-field='comma']")?.checked,
      space: document.querySelector("[data-data-field='space']")?.checked,
    }));
    await page.locator('[data-data-dialog-action="text-columns-next"]').click();
    await page.locator('[data-data-dialog-action="text-columns-next"]').click();
    const dateOptions = await page.locator("[data-data-field='dateOrder'] option").evaluateAll((options) =>
      options.map((option) => option.value),
    );
    await page.locator("[data-data-field='columnFormat'][value='date']").check();
    await page.locator("[data-data-field='dateOrder']").selectOption("MDY");
    const selectedHeader = await page.locator(".text-columns-preview-table th button.is-selected").first().innerText();
    await page.locator("[data-data-field='destination']").fill("B1");
    await page.locator('[data-data-dialog-action="text-columns-finish"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const overwritePrompt = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("#messageBoxOverlay [data-message-box-close]")].map((button) =>
        button.textContent.replace(/\s+/g, " ").trim(),
      ),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-result='cancel']").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    const afterCancel = await page.evaluate(() => ({
      dialogVisible: Boolean(document.querySelector('[data-data-dialog="text-to-columns"]')),
      b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
      c2: document.querySelector('.sheet-cell[data-address="C2"] .cell-content')?.textContent || "",
    }));
    await page.locator('[data-data-dialog-action="text-columns-finish"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    await page.locator("#messageBoxOverlay [data-message-box-result='ok']").click();
    await page.waitForSelector('[data-data-dialog="text-to-columns"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B3"] .cell-content')?.textContent === "2026/05/20",
      null,
      { timeout: 5000 },
    );
    const values = await page.evaluate(() =>
      ["B1", "C1", "D1", "B2", "C2", "D2", "B3", "C3", "D3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
    );
    const status = await page.locator("#statusText").innerText();
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "既存",
      null,
      { timeout: 5000 },
    );
    const afterUndo = await page.evaluate(() => ({
      b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
      c2: document.querySelector('.sheet-cell[data-address="C2"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    return { delimiterDefaults, dateOptions, selectedHeader, overwritePrompt, afterCancel, values, status, afterUndo };
  } finally {
    await page.close();
  }
}



async function exerciseTextToColumnsFixedWidth(context) {
  const workbookPath = await createTextToColumnsFixedWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?textColumnsFixed=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="A2"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="A2"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-text-to-columns"]').click();
    await page.waitForSelector('[data-data-dialog="text-to-columns"]', { timeout: 5000 });
    await page.locator("[data-data-field='mode'][value='fixed']").check();
    await page.locator('[data-data-dialog-action="text-columns-next"]').click();
    const inferredBreaks = await page.locator("[data-data-field='fixedBreaks']").inputValue();
    const breakMarkersBefore = await page.locator(".text-columns-step.is-active .text-columns-fixed-break.is-break").count();
    await page.locator(".text-columns-step.is-active [data-text-fixed-position='5']").click();
    const breaksAfterRemove = await page.locator("[data-data-field='fixedBreaks']").inputValue();
    await page.locator(".text-columns-step.is-active [data-text-fixed-position='5']").click();
    const breaksAfterAdd = await page.locator("[data-data-field='fixedBreaks']").inputValue();
    const previewStep2 = await page.locator("[data-text-columns-preview]").first().innerText();
    const fixedRowsText = await page.locator(".text-columns-step.is-active .text-columns-fixed-row").evaluateAll((rows) =>
      rows.map((row) => row.textContent || ""),
    );
    await page.locator('[data-data-dialog-action="text-columns-next"]').click();
    await page.locator(".text-columns-step.is-active [data-text-column-index='1']").click();
    await page.locator("[data-data-field='columnFormat'][value='text']").check();
    await page.locator(".text-columns-step.is-active [data-text-column-index='2']").click();
    await page.locator("[data-data-field='columnFormat'][value='skip']").check();
    const previewStep3 = await page.locator("[data-text-columns-preview]").first().innerText();
    await page.locator("[data-data-field='destination']").fill("B1");
    await page.locator('[data-data-dialog-action="text-columns-finish"]').click();
    await page.waitForSelector('[data-data-dialog="text-to-columns"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="C2"] .cell-content')?.textContent === "010",
      null,
      { timeout: 5000 },
    );
    const values = await page.evaluate(() =>
      ["B1", "C1", "D1", "B2", "C2", "D2"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
    );
    const status = await page.locator("#statusText").innerText();
    return { inferredBreaks, breakMarkersBefore, breaksAfterRemove, breaksAfterAdd, previewStep2, fixedRowsText, previewStep3, values, status };
  } finally {
    await page.close();
  }
}



async function exerciseTextToColumnsAdvanced(context) {
  const workbookPath = await createTextToColumnsAdvancedWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?textColumnsAdvanced=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="A2"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="A2"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-text-to-columns"]').click();
    await page.waitForSelector('[data-data-dialog="text-to-columns"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="text-columns-next"]').click();
    const delimiterDefaults = await page.evaluate(() => ({
      tab: document.querySelector("[data-data-field='tab']")?.checked,
      semicolon: document.querySelector("[data-data-field='semicolon']")?.checked,
      comma: document.querySelector("[data-data-field='comma']")?.checked,
      space: document.querySelector("[data-data-field='space']")?.checked,
    }));
    await page.locator("[data-data-field='comma']").uncheck();
    await page.locator("[data-data-field='semicolon']").check();
    await page.locator('[data-data-dialog-action="text-columns-next"]').click();
    await page.locator("[data-data-field='decimalSeparator']").fill(",");
    await page.locator("[data-data-field='thousandsSeparator']").fill(".");
    await page.locator("[data-data-field='trailingMinus']").check();
    await page.locator("[data-data-field='destination']").fill("B1");
    await page.locator('[data-data-dialog-action="text-columns-finish"]').click();
    await page.waitForSelector('[data-data-dialog="text-to-columns"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="C2"] .cell-content')?.textContent === "-45",
      null,
      { timeout: 5000 },
    );
    const display = await page.evaluate(() =>
      ["B1", "C1", "B2", "C2"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
    );
    const raw = await page.evaluate(() => ({
      b1: window.__webSheetDevCellModel?.("TextColumnsAdvanced", 1, 2)?.raw,
      c1: window.__webSheetDevCellModel?.("TextColumnsAdvanced", 1, 3)?.raw,
      b2: window.__webSheetDevCellModel?.("TextColumnsAdvanced", 2, 2)?.raw,
      c2: window.__webSheetDevCellModel?.("TextColumnsAdvanced", 2, 3)?.raw,
    }));
    const status = await page.locator("#statusText").innerText();
    return { delimiterDefaults, display, raw, status };
  } finally {
    await page.close();
  }
}



async function exerciseSortDialogLevels(context) {
  const workbookPath = await createSortDialogWorkbook();
  const sortPage = await context.newPage();
  try {
    await sortPage.goto(`${appUrl}?sortDialog=${Date.now()}`, { waitUntil: "load" });
    await sortPage.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(sortPage, workbookPath);
    await sortPage.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await sortPage.locator('.sheet-cell[data-address="A1"]').click();
    await sortPage.locator('.sheet-cell[data-address="B5"]').click({ modifiers: ["Shift"] });
    await sortPage.locator('[data-ribbon-tab="data"]').click();
    await sortPage.locator('[data-ribbon-command="data-sort-dialog"]').click();
    await sortPage.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
    const initialCount = await sortPage.locator("[data-sort-level-row]").count();
    const toolbarLabels = await sortPage.locator(".sort-dialog-toolbar button").evaluateAll((buttons) =>
      buttons.map((button) => button.textContent.replace(/\s+/g, " ").trim()),
    );
    await sortPage.locator('[data-data-dialog-action="sort-add-level"]').click();
    const afterAddCount = await sortPage.locator("[data-sort-level-row]").count();
    await sortPage.locator('[data-data-dialog-action="sort-copy-level"]').click();
    const afterCopyCount = await sortPage.locator("[data-sort-level-row]").count();
    await sortPage.locator('[data-data-dialog-action="sort-delete-level"]').click();
    const afterDeleteCount = await sortPage.locator("[data-sort-level-row]").count();
    await sortPage.selectOption('[data-data-sort-key="0"]', "1");
    await sortPage.selectOption('[data-data-sort-order="0"]', "asc");
    await sortPage.selectOption('[data-data-sort-key="1"]', "2");
    await sortPage.selectOption('[data-data-sort-order="1"]', "desc");
    await sortPage.locator('[data-data-dialog-action="ok"]').click();
    await sortPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "A",
      null,
      { timeout: 5000 },
    );
    const rows = await sortPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        group: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    await sortPage.locator('[data-ribbon-command="data-sort-dialog"]').click();
    await sortPage.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
    const reopened = await sortPage.evaluate(() => ({
      levelCount: document.querySelectorAll("[data-sort-level-row]").length,
      keys: [...document.querySelectorAll("[data-sort-level-row]")].map((row) => {
        const index = row.getAttribute("data-sort-level-row");
        return row.querySelector(`[data-data-sort-key="${index}"]`)?.value || "";
      }),
      orders: [...document.querySelectorAll("[data-sort-level-row]")].map((row) => {
        const index = row.getAttribute("data-sort-level-row");
        return row.querySelector(`[data-data-sort-order="${index}"]`)?.value || "";
      }),
      header: document.querySelector("[data-data-field='header']")?.checked || false,
      orientation: document.querySelector("[data-data-field='sortOrientation']:checked")?.value || "",
    }));
    await sortPage.locator('[data-data-dialog-action="cancel"]').click();
    return { initialCount, afterAddCount, afterCopyCount, afterDeleteCount, toolbarLabels, rows, reopened };
  } finally {
    await sortPage.close();
  }
}



async function exerciseSortDialogColorOrder(context) {
  const workbookPath = await createSortDialogWorkbook();
  const sortPage = await context.newPage();
  try {
    await sortPage.goto(`${appUrl}?sortColor=${Date.now()}`, { waitUntil: "load" });
    await sortPage.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(sortPage, workbookPath);
    await sortPage.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await sortPage.locator('.sheet-cell[data-address="A1"]').click();
    await sortPage.locator('.sheet-cell[data-address="B5"]').click({ modifiers: ["Shift"] });
    await sortPage.locator('[data-ribbon-tab="data"]').click();
    await sortPage.locator('[data-ribbon-command="data-sort-dialog"]').click();
    await sortPage.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
    await sortPage.selectOption('[data-data-sort-key="0"]', "2");
    await sortPage.selectOption('[data-data-sort-on="0"]', "cellColor");
    await sortPage.waitForSelector('[data-data-sort-color="0"]', { timeout: 5000 });
    const colorOptions = await sortPage.locator('[data-data-sort-color="0"] option').evaluateAll((options) =>
      options.map((option) => option.value),
    );
    const position = await sortPage.locator('[data-data-sort-position="0"]').inputValue();
    await sortPage.locator('[data-data-dialog-action="ok"]').click();
    await sortPage.waitForSelector('[data-data-dialog="sort"]', { state: "detached", timeout: 5000 });
    await sortPage.waitForTimeout(500);
    const rows = await sortPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        group: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const colorStatus = await sortPage.locator("#statusText").innerText();
    const colorExported = await exportedSortColorWorkbookState(
      await sortPage.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );

    await sortPage.locator('[data-ribbon-command="data-sort-dialog"]').click();
    await sortPage.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
    await sortPage.selectOption('[data-data-sort-key="0"]', "2");
    await sortPage.selectOption('[data-data-sort-on="0"]', "fontColor");
    await sortPage.waitForSelector('[data-data-sort-color="0"]', { timeout: 5000 });
    const fontColorOptions = await sortPage.locator('[data-data-sort-color="0"] option').evaluateAll((options) =>
      options.map((option) => option.value),
    );
    const fontPosition = await sortPage.locator('[data-data-sort-position="0"]').inputValue();
    await sortPage.locator('[data-data-dialog-action="ok"]').click();
    await sortPage.waitForSelector('[data-data-dialog="sort"]', { state: "detached", timeout: 5000 });
    await sortPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent === "5",
      null,
      { timeout: 5000 },
    );
    const fontRows = await sortPage.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        group: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const fontStatus = await sortPage.locator("#statusText").innerText();
    const fontExported = await exportedSortColorWorkbookState(
      await sortPage.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
    );
    return { colorOptions, position, rows, colorStatus, colorExported, fontColorOptions, fontPosition, fontRows, fontStatus, fontExported };
  } finally {
    await sortPage.close();
  }
}



async function exportedSortColorWorkbookState(base64) {
  const buffer = Buffer.from(String(base64 || ""), "base64");
  const zip = await JSZip.loadAsync(buffer);
  const [stylesXml, worksheetXmlEntries] = await Promise.all([
    zip.file("xl/styles.xml")?.async("text"),
    Promise.all(
      Object.keys(zip.files)
        .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
        .map(async (name) => [name, await zip.file(name).async("text")]),
    ),
  ]);
  const sortConditions = worksheetXmlEntries.flatMap(([, xml]) =>
    [...String(xml || "").matchAll(/<sortCondition\b([^>]*)\/>/g)].map((match) => match[1] || ""),
  );
  const colorSortConditions = sortConditions.filter((attrs) => /\bsortBy="(?:cellColor|fontColor)"/.test(attrs));
  const dxfRgbColors = [...String(stylesXml || "").matchAll(/<dxf\b[\s\S]*?<\/dxf>/g)]
    .flatMap((match) => [...match[0].matchAll(/\brgb="([0-9A-Fa-f]{8})"/g)].map((colorMatch) => colorMatch[1].toUpperCase()));
  return {
    colorSortCount: colorSortConditions.length,
    cellColorSortCount: colorSortConditions.filter((attrs) => /\bsortBy="cellColor"/.test(attrs)).length,
    fontColorSortCount: colorSortConditions.filter((attrs) => /\bsortBy="fontColor"/.test(attrs)).length,
    dxfSortCount: colorSortConditions.filter((attrs) => /\bdxfId="\d+"/.test(attrs)).length,
    inlineColorSortCount: colorSortConditions.filter((attrs) => /\bcolor=/.test(attrs)).length,
    dxfRgbColors,
  };
}



async function exerciseSortDialogLeftToRight(context) {
  const workbookPath = await createSortColumnsWorkbook();
  const sortPage = await context.newPage();
  try {
    await sortPage.goto(`${appUrl}?sortColumns=${Date.now()}`, { waitUntil: "load" });
    await sortPage.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(sortPage, workbookPath);
    await sortPage.waitForSelector('.sheet-cell[data-address="D3"]', { timeout: 10000 });
    await sortPage.locator('.sheet-cell[data-address="A1"]').click();
    await sortPage.locator('.sheet-cell[data-address="D3"]').click({ modifiers: ["Shift"] });
    await sortPage.locator('[data-ribbon-tab="data"]').click();
    await sortPage.locator('[data-ribbon-command="data-sort-dialog"]').click();
    await sortPage.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
    await sortPage.locator('[data-data-dialog-action="sort-options-toggle"]').click();
    await sortPage.locator("[data-data-field='sortOrientation'][value='columns']").check();
    await sortPage.waitForFunction(
      () => document.querySelector("[data-sort-axis-caption]")?.textContent === "行",
      null,
      { timeout: 5000 },
    );
    const optionState = await sortPage.evaluate(() => ({
      axisCaption: document.querySelector("[data-sort-axis-caption]")?.textContent || "",
      headerLabel: document.querySelector("[data-sort-header-label]")?.textContent || "",
      keys: [...document.querySelectorAll('[data-data-sort-key="0"] option')].map((option) => option.textContent.trim()),
    }));
    await sortPage.selectOption('[data-data-sort-key="0"]', "2");
    await sortPage.selectOption('[data-data-sort-order="0"]', "asc");
    await sortPage.locator('[data-data-dialog-action="ok"]').click();
    await sortPage.waitForSelector('[data-data-dialog="sort"]', { state: "detached", timeout: 5000 });
    await sortPage.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "East",
      null,
      { timeout: 5000 },
    );
    const rows = await sortPage.evaluate(() =>
      [1, 2, 3].map((row) =>
        ["A", "B", "C", "D"].map((col) => document.querySelector(`.sheet-cell[data-address="${col}${row}"] .cell-content`)?.textContent || ""),
      ),
    );
    const status = await sortPage.locator("#statusText").innerText();
    return { optionState, rows, status };
  } finally {
    await sortPage.close();
  }
}



async function exerciseSortDialogBlankAndCase(context) {
  const runBlankSort = async (direction) => {
    const workbookPath = await createSortBlankWorkbook();
    const page = await context.newPage();
    try {
      await page.goto(`${appUrl}?sortBlank=${direction}-${Date.now()}`, { waitUntil: "load" });
      await page.setInputFiles("#fileInput", workbookPath);
      await waitForUploadedWorkbook(page, workbookPath);
      await page.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
      await page.locator('.sheet-cell[data-address="A1"]').click();
      await page.locator('.sheet-cell[data-address="B5"]').click({ modifiers: ["Shift"] });
      await page.locator('[data-ribbon-tab="data"]').click();
      await page.locator('[data-ribbon-command="data-sort-dialog"]').click();
      await page.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
      await page.selectOption('[data-data-sort-key="0"]', "2");
      await page.selectOption('[data-data-sort-order="0"]', direction);
      await page.locator('[data-data-dialog-action="ok"]').click();
      await page.waitForSelector('[data-data-dialog="sort"]', { state: "detached", timeout: 5000 });
      await page.waitForFunction(
        (expectedFirst) => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === expectedFirst,
        direction === "desc" ? "High" : "Low",
        { timeout: 5000 },
      );
      const rows = await page.evaluate(() =>
        [2, 3, 4, 5].map((row) => ({
          row,
          name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
          score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
        })),
      );
      const status = await page.locator("#statusText").innerText();
      return { rows, status };
    } finally {
      await page.close();
    }
  };

  const runCaseSensitiveSort = async () => {
    const workbookPath = await createSortCaseWorkbook();
    const page = await context.newPage();
    try {
      await page.goto(`${appUrl}?sortCase=${Date.now()}`, { waitUntil: "load" });
      await page.setInputFiles("#fileInput", workbookPath);
      await waitForUploadedWorkbook(page, workbookPath);
      await page.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
      await page.locator('.sheet-cell[data-address="A1"]').click();
      await page.locator('.sheet-cell[data-address="B5"]').click({ modifiers: ["Shift"] });
      await page.locator('[data-ribbon-tab="data"]').click();
      await page.locator('[data-ribbon-command="data-sort-dialog"]').click();
      await page.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
      await page.locator('[data-data-dialog-action="sort-options-toggle"]').click();
      await page.locator("[data-data-field='caseSensitive']").check();
      await page.selectOption('[data-data-sort-key="0"]', "1");
      await page.selectOption('[data-data-sort-order="0"]', "asc");
      await page.locator('[data-data-dialog-action="ok"]').click();
      await page.waitForSelector('[data-data-dialog="sort"]', { state: "detached", timeout: 5000 });
      await page.waitForFunction(
        () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "a",
        null,
        { timeout: 5000 },
      );
      const rows = await page.evaluate(() =>
        [2, 3, 4, 5].map((row) => ({
          key: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
          value: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
        })),
      );
      const status = await page.locator("#statusText").innerText();
      return { rows, status };
    } finally {
      await page.close();
    }
  };

  return {
    blankAsc: await runBlankSort("asc"),
    blankDesc: await runBlankSort("desc"),
    caseSensitive: await runCaseSensitiveSort(),
  };
}



async function exerciseSortDialogCustomList(context) {
  const runCustomSort = async ({ keyCol, listId, expectedFirst }) => {
    const workbookPath = await createSortCustomListWorkbook();
    const page = await context.newPage();
    try {
      await page.goto(`${appUrl}?sortCustomList=${listId}-${Date.now()}`, { waitUntil: "load" });
      await page.setInputFiles("#fileInput", workbookPath);
      await waitForUploadedWorkbook(page, workbookPath);
      await page.waitForSelector('.sheet-cell[data-address="D5"]', { timeout: 10000 });
      await page.locator('.sheet-cell[data-address="A1"]').click();
      await page.locator('.sheet-cell[data-address="D5"]').click({ modifiers: ["Shift"] });
      await page.locator('[data-ribbon-tab="data"]').click();
      await page.locator('[data-ribbon-command="data-sort-dialog"]').click();
      await page.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
      const orderOptions = await page.locator('[data-data-sort-order="0"] option').evaluateAll((options) =>
        options.map((option) => ({ value: option.value, label: option.textContent.replace(/\s+/g, " ").trim() })),
      );
      await page.selectOption('[data-data-sort-key="0"]', String(keyCol));
      await page.selectOption('[data-data-sort-order="0"]', `custom:${listId}`);
      await page.locator('[data-data-dialog-action="ok"]').click();
      await page.waitForSelector('[data-data-dialog="sort"]', { state: "detached", timeout: 5000 });
      await page.waitForFunction(
        (first) => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === first,
        expectedFirst,
        { timeout: 5000 },
      );
      const rows = await page.evaluate(() =>
        [2, 3, 4, 5].map((row) => ({
          month: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
          value: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
          weekday: document.querySelector(`.sheet-cell[data-address="C${row}"] .cell-content`)?.textContent || "",
          code: document.querySelector(`.sheet-cell[data-address="D${row}"] .cell-content`)?.textContent || "",
        })),
      );
      const status = await page.locator("#statusText").innerText();
      const exported = await exportedCustomListSortWorkbookState(
        await page.evaluate(() => window.__webSheetDevExportXlsxBase64?.() || ""),
      );
      return { orderOptions, rows, status, exported };
    } finally {
      await page.close();
    }
  };
  return {
    months: await runCustomSort({ keyCol: 1, listId: "month-ja", expectedFirst: "1月" }),
    weekdays: await runCustomSort({ keyCol: 3, listId: "weekday-ja", expectedFirst: "1月" }),
  };
}



async function exportedCustomListSortWorkbookState(base64) {
  const buffer = Buffer.from(String(base64 || ""), "base64");
  const zip = await JSZip.loadAsync(buffer);
  const worksheetXmlEntries = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
      .map(async (name) => [name, await zip.file(name).async("text")]),
  );
  const customListTexts = worksheetXmlEntries.flatMap(([, xml]) =>
    [...String(xml || "").matchAll(/<sortCondition\b([^>]*)\/>/g)]
      .map((match) => String(match[1] || "").match(/\bcustomList="([^"]*)"/)?.[1] || "")
      .filter(Boolean),
  );
  return {
    customListSortCount: customListTexts.length,
    customListTexts,
  };
}



async function exerciseSortDialogDateOrder(context) {
  const workbookPath = await createSortDateWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?sortDate=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B6"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-sort-dialog"]').click();
    await page.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
    await page.selectOption('[data-data-sort-key="0"]', "1");
    await page.selectOption('[data-data-sort-order="0"]', "asc");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="sort"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent === "Prev",
      null,
      { timeout: 5000 },
    );
    const rows = await page.evaluate(() =>
      [2, 3, 4, 5, 6].map((row) => ({
        date: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        label: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const status = await page.locator("#statusText").innerText();

    await page.locator('.sheet-cell[data-address="C1"]').click();
    await page.locator('.sheet-cell[data-address="D5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-sort-dialog"]').click();
    await page.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
    await page.selectOption('[data-data-sort-key="0"]', "3");
    await page.selectOption('[data-data-sort-order="0"]', "asc");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="sort"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="D2"] .cell-content')?.textContent === "Prev",
      null,
      { timeout: 5000 },
    );
    const rawDateRows = await page.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        date: document.querySelector(`.sheet-cell[data-address="C${row}"] .cell-content`)?.textContent || "",
        label: document.querySelector(`.sheet-cell[data-address="D${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const rawDateStatus = await page.locator("#statusText").innerText();
    return { rows, status, rawDateRows, rawDateStatus };
  } finally {
    await page.close();
  }
}



async function exerciseSortFormulaResults(context) {
  const workbookPath = await createSortFormulaWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?sortFormula=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-sort-dialog"]').click();
    await page.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
    await page.selectOption('[data-data-sort-key="0"]', "2");
    await page.selectOption('[data-data-sort-order="0"]', "asc");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="sort"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "One",
      null,
      { timeout: 5000 },
    );
    const rows = await page.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        calc: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
        formula: window.__webSheetDebug?.getCell?.(row, 2)?.formula || "",
      })),
    );
    const status = await page.locator("#statusText").innerText();
    await page.locator('.sheet-cell[data-address="B4"]').click();
    await page.locator('[data-ribbon-command="data-sort-desc"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A2"] .cell-content')?.textContent === "TenA",
      null,
      { timeout: 5000 },
    );
    const quickRows = await page.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        calc: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const quickStatus = await page.locator("#statusText").innerText();
    return { rows, status, quickRows, quickStatus };
  } finally {
    await page.close();
  }
}



async function exerciseQuickSortActiveColumn(context) {
  const workbookPath = await createSortDialogWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?quickSort=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="B3"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-sort-desc"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent === "20",
      null,
      { timeout: 5000 },
    );
    const rows = await page.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        group: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const selectionLabel = await page.locator("#selectedCellLabel").innerText();
    const status = await page.locator("#statusText").innerText();
    return { rows, selectionLabel, status };
  } finally {
    await page.close();
  }
}



async function exerciseQuickSortPartialSelection(context) {
  const workbookPath = await createSortDialogWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?quickSortPartial=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="B2"]').click();
    await page.locator('.sheet-cell[data-address="B5"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-sort-desc"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent === "20",
      null,
      { timeout: 5000 },
    );
    const rows = await page.evaluate(() =>
      [2, 3, 4, 5].map((row) => ({
        row,
        group: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    const status = await page.locator("#statusText").innerText();
    return { rows, status };
  } finally {
    await page.close();
  }
}



async function exerciseRemoveDuplicatesDialog(context) {
  const workbookPath = await createRemoveDuplicatesWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?removeDuplicates=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B6"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('.sheet-cell[data-address="B6"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-remove-duplicates"]').click();
    await page.waitForSelector('[data-data-dialog="remove-duplicates"]', { timeout: 5000 });
    const labelsWithHeader = await page.locator("[data-duplicate-col-label]").evaluateAll((labels) => labels.map((label) => label.textContent.trim()));
    await page.locator("[data-data-field='header']").uncheck();
    const labelsWithoutHeader = await page.locator("[data-duplicate-col-label]").evaluateAll((labels) => labels.map((label) => label.textContent.trim()));
    await page.locator("[data-data-field='header']").check();
    await page.locator('[data-data-dialog-action="unselect-all-columns"]').click();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const noColumnMessage = await page.locator("#messageBoxText").innerText();
    const noColumnTitle = await page.locator("#messageBoxTitle").innerText();
    const noColumnDialogStillOpen = await page.locator('[data-data-dialog="remove-duplicates"]').isVisible();
    const uncheckedCount = await page.locator("[data-data-duplicate-col]:not(:checked)").count();
    await page.locator("#messageBoxOverlay [data-message-box-close]").first().click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await page.locator('[data-data-dialog-action="select-all-columns"]').click();
    const recheckedCount = await page.locator("[data-data-duplicate-col]:checked").count();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector("#messageBox", { timeout: 5000 });
    const message = await page.locator("#messageBoxText").innerText();
    const title = await page.locator("#messageBoxTitle").innerText();
    const rows = await page.evaluate(() =>
      [2, 3, 4, 5, 6].map((row) => ({
        row,
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
    );
    await page.locator("#messageBoxOverlay [data-message-box-close]").first().click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="A6"] .cell-content')?.textContent === "a",
      null,
      { timeout: 5000 },
    );
    const afterUndo = await page.evaluate(() => ({
      rows: [2, 3, 4, 5, 6].map((row) => ({
        row,
        name: document.querySelector(`.sheet-cell[data-address="A${row}"] .cell-content`)?.textContent || "",
        score: document.querySelector(`.sheet-cell[data-address="B${row}"] .cell-content`)?.textContent || "",
      })),
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('.sheet-cell[data-address="D1"]').click();
    await page.locator('.sheet-cell[data-address="E4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-remove-duplicates"]').click();
    await page.waitForSelector('[data-data-dialog="remove-duplicates"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector("#messageBox", { timeout: 5000 });
    const formulaDisplayDuplicates = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      rows: [2, 3, 4].map((row) => ({
        calc: document.querySelector(`.sheet-cell[data-address="D${row}"] .cell-content`)?.textContent || "",
        group: document.querySelector(`.sheet-cell[data-address="E${row}"] .cell-content`)?.textContent || "",
      })),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator("#messageBoxOverlay [data-message-box-close]").first().click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });

    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="D3"] .cell-content')?.textContent === "2",
      null,
      { timeout: 5000 },
    );
    await page.locator('.sheet-cell[data-address="F1"]').click();
    await page.locator('.sheet-cell[data-address="G4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-remove-duplicates"]').click();
    await page.waitForSelector('[data-data-dialog="remove-duplicates"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector("#messageBox", { timeout: 5000 });
    const dateValueDuplicates = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      rows: [2, 3, 4].map((row) => ({
        date: document.querySelector(`.sheet-cell[data-address="F${row}"] .cell-content`)?.textContent || "",
        group: document.querySelector(`.sheet-cell[data-address="G${row}"] .cell-content`)?.textContent || "",
      })),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator("#messageBoxOverlay [data-message-box-close]").first().click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    return { labelsWithHeader, labelsWithoutHeader, noColumnTitle, noColumnMessage, noColumnDialogStillOpen, uncheckedCount, recheckedCount, title, message, rows, afterUndo, formulaDisplayDuplicates, dateValueDuplicates };
  } finally {
    await page.close();
  }
}

export {
  exerciseConsolidateDialog,
  exerciseDataAnalysisToolpak,
  exerciseDataGetConnectionMenus,
  exerciseForecastSheet,
  exerciseOutlineGrouping,
  exerciseQuickSortActiveColumn,
  exerciseQuickSortPartialSelection,
  exerciseRefreshAllUndo,
  exerciseRelationshipsDialog,
  exerciseRemoveDuplicatesDialog,
  exerciseSortDialogBlankAndCase,
  exerciseSortDialogColorOrder,
  exerciseSortDialogCustomList,
  exerciseSortDialogDateOrder,
  exerciseSortDialogLeftToRight,
  exerciseSortDialogLevels,
  exerciseSortFormulaResults,
  exerciseSubtotalDialog,
  exerciseTextToColumnsAdvanced,
  exerciseTextToColumnsDialog,
  exerciseTextToColumnsFixedWidth,
  exportedCustomListSortWorkbookState,
  exportedSortColorWorkbookState,
};
