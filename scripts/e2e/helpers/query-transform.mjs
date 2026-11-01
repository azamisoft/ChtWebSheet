// Sequential E2E feature module.
import {
  appUrl,
  defaultWorkbook,
  workbookPath,
} from "../../e2e-real-upload.mjs";
import {
  createAutoFilterMenuWorkbook,
  createFlashFillWorkbook,
  createPowerQueryTransformWorkbook,
  createWhatIfWorkbook,
  waitForUploadedWorkbook,
} from "./fixtures-and-model.mjs";


async function exercisePowerQueryTransformEditor(context) {
  const workbookPath = await createPowerQueryTransformWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?powerQueryTransform=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="G6"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A1:C6");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    const tableInitial = await page.evaluate(() => ({
      title: document.querySelector("#queryTransformTitle")?.textContent || "",
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      source: document.querySelector(".power-query-formula-bar")?.textContent || "",
    }));
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "City"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("City");
    await page.locator("[data-data-field='queryRenameTo']").fill("Area");
    await page.locator('[data-data-dialog-action="query-transform-rename-column"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Area"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Score");
    await page.locator("[data-data-field='queryDataType']").selectOption("number");
    await page.locator('[data-data-dialog-action="query-transform-change-type"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Score");
    await page.locator('[data-data-dialog-action="query-transform-sort-desc"]').click();
    await page.locator('[data-data-dialog-action="query-transform-remove-duplicates"]').click();
    await page.locator('[data-data-dialog-action="query-transform-remove-blank-rows"]').click();
    await page.locator("[data-data-field='queryRowCount']").fill("2");
    await page.locator('[data-data-dialog-action="query-transform-keep-top-rows"]').click();
    const tableAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("E1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="E2"] .cell-content')?.textContent === "Beta" &&
        document.querySelector('.sheet-cell[data-address="G3"] .cell-content')?.textContent === "Tokyo",
      null,
      { timeout: 5000 },
    );
    const tableLoadTo = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || []).find((entry) => entry.query?.sourceType === "tableRange")?.query || null;
      return {
        values: ["E1", "F1", "G1", "E2", "F2", "G2", "E3", "F3", "G3"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawScore: window.__webSheetDevCellModel?.("PQ", 2, 6)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A1:C6");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "City"));
    await page.locator('[data-data-dialog-action="query-transform-keep-duplicates"]').click();
    const keepDuplicatesAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("BY1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => window.__webSheetDevCellModel?.("PQ", 2, 77)?.raw === "Alpha" &&
        window.__webSheetDevCellModel?.("PQ", 3, 79)?.raw === "Tokyo",
      null,
      { timeout: 5000 },
    );
    const keepDuplicatesLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .filter((entry) => entry.query?.sourceType === "tableRange")
        .find((entry) => entry.query?.range?.left === 77)?.query || null;
      const cellText = (row, col) => {
        const model = window.__webSheetDevCellModel?.("PQ", row, col);
        if (!model) return "";
        return String(model.display ?? model.raw ?? "");
      };
      return {
        values: [
          cellText(1, 77), cellText(1, 78), cellText(1, 79),
          cellText(2, 77), cellText(2, 78), cellText(2, 79),
          cellText(3, 77), cellText(3, 78), cellText(3, 79),
        ],
        query,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
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

    await page.locator('.sheet-cell[data-address="A1"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A1:C6");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "City"));
    await page.locator('[data-data-dialog-action="query-transform-keep-blank-rows"]').click();
    const keepBlankRowsAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("CC1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => window.__webSheetDevCellModel?.("PQ", 1, 81)?.raw === "Name" &&
        window.__webSheetDevQueries?.().some((entry) => entry.query?.range?.left === 81),
      null,
      { timeout: 5000 },
    );
    const keepBlankRowsLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .filter((entry) => entry.query?.sourceType === "tableRange")
        .find((entry) => entry.query?.range?.left === 81)?.query || null;
      const cellText = (row, col) => {
        const model = window.__webSheetDevCellModel?.("PQ", row, col);
        if (!model) return "";
        return String(model.display ?? model.raw ?? "");
      };
      return {
        values: [
          cellText(1, 81), cellText(1, 82), cellText(1, 83),
          cellText(2, 81), cellText(2, 82), cellText(2, 83),
        ],
        query,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
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
    await page.waitForSelector('.sheet-cell[data-address="J1"]', { timeout: 10000 });

    await page.locator('.sheet-cell[data-address="J1"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="get-web"]').click();
    await page.waitForSelector('[data-data-dialog="web-import"]', { timeout: 5000 });
    const webDataUrl = "data:text/csv;charset=utf-8,Name%2CScore%2CCity%0AWebAlpha%2C9%2CTokyo%0AWebBeta%2C7%2COsaka%0AWebGamma%2C8%2CTokyo";
    await page.locator("[data-data-field='sourceUrl']").fill(webDataUrl);
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Score"));
    await page.locator("[data-query-transform-column][value='Score']").check();
    await page.locator('[data-data-dialog-action="query-transform-remove-columns"]').click();
    await page.locator("[data-data-field='queryFilterColumn']").selectOption("City");
    await page.locator("[data-data-field='queryFilterText']").fill("Tokyo");
    await page.locator('[data-data-dialog-action="query-transform-apply-filter"]').click();
    const webAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load"]').first().click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="J2"] .cell-content')?.textContent === "WebAlpha" &&
        document.querySelector('.sheet-cell[data-address="K3"] .cell-content')?.textContent === "Tokyo",
      null,
      { timeout: 5000 },
    );
    const webLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || []).find((entry) => entry.query?.sourceType === "web")?.query || null;
      return {
        values: ["J1", "K1", "J2", "K2", "J3", "K3"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="get-web"]').click();
    await page.waitForSelector('[data-data-dialog="web-import"]', { timeout: 5000 });
    const webNavigatorHtml = `
      <html><body>
        <h2>Summary</h2>
        <table><tr><th>Metric</th><th>Total</th></tr><tr><td>Count</td><td>3</td></tr></table>
        <h2>Work Items</h2>
        <table id="work-items">
          <tr><th>Task</th><th>Owner</th><th>Status</th></tr>
          <tr><td>Build</td><td>Mika</td><td>Open</td></tr>
          <tr><td>Ship</td><td>Ken</td><td>Closed</td></tr>
          <tr><td>Review</td><td>Ren</td><td>Open</td></tr>
        </table>
      </body></html>`;
    await page.locator("[data-data-field='sourceUrl']").fill(`data:text/html;charset=utf-8,${encodeURIComponent(webNavigatorHtml)}`);
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="web-navigator"]', { timeout: 5000 });
    const webNavigatorInitial = await page.evaluate(() => ({
      sources: [...document.querySelectorAll(".web-navigator-list button span")].map((item) => item.textContent.trim()),
      active: document.querySelector(".web-navigator-list button.is-active span")?.textContent.trim() || "",
      preview: document.querySelector(".web-navigator-preview")?.textContent.replace(/\s+/g, "") || "",
      summary: document.querySelector("[data-web-navigator-summary]")?.textContent.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-web-source-index="1"]').click();
    const webNavigatorSelected = await page.evaluate(() => ({
      active: document.querySelector(".web-navigator-list button.is-active span")?.textContent.trim() || "",
      preview: document.querySelector(".web-navigator-preview")?.textContent.replace(/\s+/g, "") || "",
      summary: document.querySelector("[data-web-navigator-summary]")?.textContent.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Status"));
    await page.locator("[data-data-field='queryFilterColumn']").selectOption("Status");
    await page.locator("[data-data-field='queryFilterText']").fill("Open");
    await page.locator('[data-data-dialog-action="query-transform-apply-filter"]').click();
    const webNavigatorAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("BM1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => window.__webSheetDevCellModel?.("PQ", 2, 65)?.raw === "Build" &&
        window.__webSheetDevCellModel?.("PQ", 3, 67)?.raw === "Open",
      null,
      { timeout: 5000 },
    );
    const webNavigatorLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .filter((entry) => entry.query?.sourceType === "web")
        .find((entry) => entry.query?.range?.left === 65)?.query || null;
      const cellText = (row, col) => {
        const model = window.__webSheetDevCellModel?.("PQ", row, col);
        if (!model) return "";
        return String(model.display ?? model.raw ?? "");
      };
      return {
        values: [
          cellText(1, 65), cellText(1, 66), cellText(1, 67),
          cellText(2, 65), cellText(2, 66), cellText(2, 67),
          cellText(3, 65), cellText(3, 66), cellText(3, 67),
        ],
        query,
        rawTask: window.__webSheetDevCellModel?.("PQ", 2, 65)?.raw,
        rawStatus: window.__webSheetDevCellModel?.("PQ", 3, 67)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="get-web"]').click();
    await page.waitForSelector('[data-data-dialog="web-import"]', { timeout: 5000 });
    const jsonNavigatorSource = {
      meta: [{ Metric: "Total", Value: 3 }],
      items: [
        { Task: "Build", Owner: "Mika", Status: "Open", Detail: { Score: 10 } },
        { Task: "Ship", Owner: "Ken", Status: "Closed", Detail: { Score: 20 } },
        { Task: "Review", Owner: "Ren", Status: "Open", Detail: { Score: 30 } },
      ],
    };
    await page.locator("[data-data-field='sourceUrl']").fill(`data:application/json;charset=utf-8,${encodeURIComponent(JSON.stringify(jsonNavigatorSource))}`);
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="web-navigator"]', { timeout: 5000 });
    const jsonNavigatorInitial = await page.evaluate(() => ({
      sources: [...document.querySelectorAll(".web-navigator-list button span")].map((item) => item.textContent.trim()),
      active: document.querySelector(".web-navigator-list button.is-active span")?.textContent.trim() || "",
      preview: document.querySelector(".web-navigator-preview")?.textContent.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-web-source-index="1"]').click();
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Status"));
    await page.locator("[data-data-field='queryFilterColumn']").selectOption("Status");
    await page.locator("[data-data-field='queryFilterText']").fill("Open");
    await page.locator('[data-data-dialog-action="query-transform-apply-filter"]').click();
    const jsonNavigatorAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("BP1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => window.__webSheetDevCellModel?.("PQ", 2, 68)?.raw === "Build" &&
        window.__webSheetDevCellModel?.("PQ", 3, 71)?.raw === 30,
      null,
      { timeout: 5000 },
    );
    const jsonNavigatorLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .filter((entry) => entry.query?.sourceType === "web")
        .find((entry) => entry.query?.range?.left === 68)?.query || null;
      const cellText = (row, col) => {
        const model = window.__webSheetDevCellModel?.("PQ", row, col);
        if (!model) return "";
        return String(model.display ?? model.raw ?? "");
      };
      return {
        values: [
          cellText(1, 68), cellText(1, 69), cellText(1, 70), cellText(1, 71),
          cellText(2, 68), cellText(2, 69), cellText(2, 70), cellText(2, 71),
          cellText(3, 68), cellText(3, 69), cellText(3, 70), cellText(3, 71),
        ],
        query,
        rawScore: window.__webSheetDevCellModel?.("PQ", 3, 71)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-menu"]').click();
    await page.waitForSelector("#dataCommandMenu", { timeout: 5000 });
    await page.locator('#dataCommandMenu [data-data-menu-action="get-web"]').click();
    await page.waitForSelector('[data-data-dialog="web-import"]', { timeout: 5000 });
    const xmlSource = `<catalog>
      <item id="A1"><Name>Alpha</Name><Qty>5</Qty><City>Tokyo</City></item>
      <item id="B2"><Name>Beta</Name><Qty>9</Qty><City>Osaka</City></item>
    </catalog>`;
    await page.locator("[data-data-field='sourceUrl']").fill(`data:application/xml;charset=utf-8,${encodeURIComponent(xmlSource)}`);
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    const xmlInitial = await page.evaluate(() => ({
      title: document.querySelector("#queryTransformTitle")?.textContent || "",
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
    }));
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Qty"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Qty");
    await page.locator("[data-data-field='queryDataType']").selectOption("number");
    await page.locator('[data-data-dialog-action="query-transform-change-type"]').click();
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("BU1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => window.__webSheetDevCellModel?.("PQ", 2, 74)?.raw === "Alpha" &&
        window.__webSheetDevCellModel?.("PQ", 3, 75)?.raw === 9,
      null,
      { timeout: 5000 },
    );
    const xmlLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .filter((entry) => entry.query?.sourceType === "web")
        .find((entry) => entry.query?.range?.left === 73)?.query || null;
      const cellText = (row, col) => {
        const model = window.__webSheetDevCellModel?.("PQ", row, col);
        if (!model) return "";
        return String(model.display ?? model.raw ?? "");
      };
      return {
        values: [
          cellText(1, 73), cellText(1, 74), cellText(1, 75), cellText(1, 76),
          cellText(2, 73), cellText(2, 74), cellText(2, 75), cellText(2, 76),
          cellText(3, 73), cellText(3, 74), cellText(3, 75), cellText(3, 76),
        ],
        query,
        rawQty: window.__webSheetDevCellModel?.("PQ", 3, 75)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
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

    await page.waitForSelector('.sheet-cell[data-address="C12"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A8"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A8:C12");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "FullName"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("FullName");
    await page.locator("[data-data-field='querySplitDelimiter']").fill("|");
    await page.locator("[data-data-field='queryReplaceFrom']").fill(" - ");
    await page.locator("[data-data-field='queryReplaceTo']").fill("|");
    await page.locator('[data-data-dialog-action="query-transform-trim"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("FullName");
    await page.locator('[data-data-dialog-action="query-transform-clean"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("FullName");
    await page.locator('[data-data-dialog-action="query-transform-replace-values"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("FullName");
    await page.locator('[data-data-dialog-action="query-transform-split-column"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "FullName.2"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("FullName");
    await page.locator('[data-data-dialog-action="query-transform-fill-down"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Amount");
    await page.locator("[data-data-field='queryDataType']").selectOption("number");
    await page.locator('[data-data-dialog-action="query-transform-change-type"]').click();
    await page.locator("[data-data-field='queryCustomColumnName']").fill("Label");
    await page.locator("[data-data-field='queryCustomFormula']").fill('[FullName] & ":" & [Category]');
    await page.locator('[data-data-dialog-action="query-transform-add-custom-column"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Label"));
    await page.locator("[data-data-field='queryGroupByColumn']").selectOption("Category");
    await page.locator("[data-data-field='queryAggregateColumn']").selectOption("Amount");
    await page.locator("[data-data-field='queryAggregateFunction']").selectOption("sum");
    await page.locator("[data-data-field='queryAggregateName']").fill("Total");
    await page.locator('[data-data-dialog-action="query-transform-group-by"]').click();
    const textAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("M1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="M2"] .cell-content')?.textContent === "A" &&
        document.querySelector('.sheet-cell[data-address="N3"] .cell-content')?.textContent === "7",
      null,
      { timeout: 5000 },
    );
    const textLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || []).find((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 8)?.query || null;
      return {
        values: ["M1", "N1", "M2", "N2", "M3", "N3"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawTotal: window.__webSheetDevCellModel?.("PQ", 2, 14)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.waitForSelector('.sheet-cell[data-address="D19"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A15"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A15:D19");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Code"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Code");
    await page.locator("[data-data-field='queryExtractDelimiter']").fill("-");
    await page.locator('[data-data-dialog-action="query-transform-extract-before"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Region");
    await page.locator('[data-data-dialog-action="query-transform-duplicate-column"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Region - コピー"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Region - コピー");
    await page.locator('[data-data-dialog-action="query-transform-move-start"]').click();
    await page.locator("[data-data-field='queryIndexStart']").fill("1");
    await page.locator('[data-data-dialog-action="query-transform-add-index-column"]').click();
    await page.locator("[data-data-field='queryConditionName']").fill("Size");
    await page.locator("[data-data-field='queryConditionColumn']").selectOption("Qty");
    await page.locator("[data-data-field='queryConditionOperator']").selectOption("greaterOrEqual");
    await page.locator("[data-data-field='queryConditionValue']").fill("10");
    await page.locator("[data-data-field='queryConditionTrueValue']").fill("大");
    await page.locator("[data-data-field='queryConditionFalseValue']").fill("小");
    await page.locator('[data-data-dialog-action="query-transform-add-conditional-column"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Size"));
    await page.locator("[data-data-field='queryFilterColumn']").selectOption("Qty");
    await page.locator("[data-data-field='queryFilterOperator']").selectOption("greaterOrEqual");
    await page.locator("[data-data-field='queryFilterText']").fill("8");
    await page.locator('[data-data-dialog-action="query-transform-apply-filter"]').click();
    await page.locator('[data-query-transform-column][value="Region - コピー"]').check();
    await page.locator('[data-query-transform-column][value="Item"]').check();
    await page.locator('[data-query-transform-column][value="Code"]').check();
    await page.locator('[data-query-transform-column][value="インデックス"]').check();
    await page.locator('[data-query-transform-column][value="Size"]').check();
    await page.locator('[data-data-dialog-action="query-transform-keep-columns"]').click();
    const columnAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("P1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="P2"] .cell-content')?.textContent === "Osaka" &&
        document.querySelector('.sheet-cell[data-address="T4"] .cell-content')?.textContent === "大",
      null,
      { timeout: 5000 },
    );
    const columnLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || []).find((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 15)?.query || null;
      return {
        values: ["P1", "Q1", "R1", "S1", "T1", "P2", "Q2", "R2", "S2", "T2", "P3", "Q3", "R3", "S3", "T3", "P4", "Q4", "R4", "S4", "T4"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawIndex: window.__webSheetDevCellModel?.("PQ", 2, 19)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.waitForSelector('.sheet-cell[data-address="A22"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="A22"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A22:E26");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Q2"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("City");
    await page.locator('[data-data-dialog-action="query-transform-uppercase"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Code");
    await page.locator("[data-data-field='queryExtractCount']").fill("3");
    await page.locator('[data-data-dialog-action="query-transform-extract-last"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Code");
    await page.locator('[data-data-dialog-action="query-transform-text-length"]').click();
    await page.locator('[data-query-transform-column][value="Q1"]').check();
    await page.locator('[data-query-transform-column][value="Q2"]').check();
    await page.locator('[data-data-dialog-action="query-transform-unpivot-columns"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "属性"));
    await page.locator('[data-query-transform-column][value="Dept"]').check();
    await page.locator('[data-query-transform-column][value="属性"]').check();
    await page.locator("[data-data-field='queryMergeDelimiter']").fill(":");
    await page.locator("[data-data-field='queryMergeName']").fill("Key");
    await page.locator('[data-data-dialog-action="query-transform-merge-columns"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Key"));
    await page.locator('[data-data-dialog-action="query-transform-reverse-rows"]').click();
    await page.locator("[data-data-field='queryAlternateStart']").fill("2");
    await page.locator("[data-data-field='queryAlternateRemove']").fill("1");
    await page.locator("[data-data-field='queryAlternateKeep']").fill("1");
    await page.locator('[data-data-dialog-action="query-transform-remove-alternate-rows"]').click();
    const reshapeAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("V1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="V2"] .cell-content')?.textContent === "South:Q2" &&
        document.querySelector('.sheet-cell[data-address="Y5"] .cell-content')?.textContent === "20",
      null,
      { timeout: 5000 },
    );
    const reshapeLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || []).find((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 22)?.query || null;
      return {
        values: ["V1", "W1", "X1", "Y1", "V2", "W2", "X2", "Y2", "V3", "W3", "X3", "Y3", "V4", "W4", "X4", "Y4", "V5", "W5", "X5", "Y5"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawCodeLength: window.__webSheetDevCellModel?.("PQ", 2, 24)?.raw,
        rawValue: window.__webSheetDevCellModel?.("PQ", 2, 25)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A29:D32");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Amount"));
    await page.locator('[data-data-dialog-action="query-transform-remove-empty-columns"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Err");
    await page.locator("[data-data-field='queryErrorReplacement']").fill("修正");
    await page.locator('[data-data-dialog-action="query-transform-replace-errors"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Amount");
    await page.locator("[data-data-field='queryNumberOperand']").fill("2");
    await page.locator('[data-data-dialog-action="query-transform-number-multiply"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Amount");
    await page.locator('[data-data-dialog-action="query-transform-number-round"]').click();
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Date");
    await page.locator('[data-data-dialog-action="query-transform-date-year"]').click();
    const numericAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("AA1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="AA2"] .cell-content')?.textContent === "2026" &&
        document.querySelector('.sheet-cell[data-address="AC4"] .cell-content')?.textContent === "修正",
      null,
      { timeout: 5000 },
    );
    const numericLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || []).find((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 29)?.query || null;
      return {
        values: ["AA1", "AB1", "AC1", "AA2", "AB2", "AC2", "AA3", "AB3", "AC3", "AA4", "AB4", "AC4"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawYear: window.__webSheetDevCellModel?.("PQ", 2, 27)?.raw,
        rawAmount: window.__webSheetDevCellModel?.("PQ", 2, 28)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A35:C37");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Metric"));
    await page.locator('[data-data-dialog-action="query-transform-demote-headers"]').click();
    await page.locator('[data-data-dialog-action="query-transform-transpose"]').click();
    const transposeAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("AE1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="AE2"] .cell-content')?.textContent === "Column1" &&
        document.querySelector('.sheet-cell[data-address="AH4"] .cell-content')?.textContent === "4",
      null,
      { timeout: 5000 },
    );
    const transposeLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || []).find((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 35)?.query || null;
      return {
        values: ["AE1", "AF1", "AG1", "AH1", "AE2", "AF2", "AG2", "AH2", "AE3", "AF3", "AG3", "AH3", "AE4", "AF4", "AG4", "AH4"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A40:C45");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "RowID"));
    await page.locator("[data-data-field='queryRowCount']").fill("1");
    await page.locator('[data-data-dialog-action="query-transform-remove-top-rows"]').click();
    await page.locator("[data-data-field='queryRowCount']").fill("1");
    await page.locator('[data-data-dialog-action="query-transform-remove-bottom-rows"]').click();
    await page.locator("[data-data-field='queryRangeStart']").fill("2");
    await page.locator("[data-data-field='queryRangeCount']").fill("2");
    await page.locator('[data-data-dialog-action="query-transform-keep-range-rows"]').click();
    await page.locator("[data-data-field='queryRowCount']").fill("1");
    await page.locator('[data-data-dialog-action="query-transform-keep-bottom-rows"]').click();
    const rowAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("AJ1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="AJ2"] .cell-content')?.textContent === "4" &&
        document.querySelector('.sheet-cell[data-address="AL2"] .cell-content')?.textContent === "tail",
      null,
      { timeout: 5000 },
    );
    const rowLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || []).find((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 40)?.query || null;
      return {
        values: ["AJ1", "AK1", "AL1", "AJ2", "AK2", "AL2"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawRowId: window.__webSheetDevCellModel?.("PQ", 2, 36)?.raw,
        rawValue: window.__webSheetDevCellModel?.("PQ", 2, 37)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A29:D32");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Err"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Err");
    await page.locator('[data-data-dialog-action="query-transform-keep-errors"]').click();
    const keepErrorsAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("AM1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="AO2"] .cell-content')?.textContent === "#N/A" &&
        document.querySelector('.sheet-cell[data-address="AO3"] .cell-content')?.textContent === "#DIV/0!",
      null,
      { timeout: 5000 },
    );
    const keepErrorsLoad = await page.evaluate(() => {
      const queries = (window.__webSheetDevQueries?.() || [])
        .filter((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 29);
      const query = queries.find((entry) => entry.query?.range?.left === 39)?.query || null;
      return {
        values: ["AM1", "AN1", "AO1", "AP1", "AM2", "AN2", "AO2", "AP2", "AM3", "AN3", "AO3", "AP3"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawError: window.__webSheetDevCellModel?.("PQ", 2, 41)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A48:C54");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Plan"));
    await page.locator("[data-data-field='queryPivotColumn']").selectOption("Plan");
    await page.locator("[data-data-field='queryPivotValueColumn']").selectOption("Amount");
    await page.locator("[data-data-field='queryPivotAggregateFunction']").selectOption("sum");
    await page.locator('[data-data-dialog-action="query-transform-pivot-column"]').click();
    const pivotAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("AQ1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="AQ2"] .cell-content')?.textContent === "East" &&
        document.querySelector('.sheet-cell[data-address="AS4"] .cell-content')?.textContent === "9",
      null,
      { timeout: 5000 },
    );
    const pivotLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 48)?.query || null;
      return {
        values: ["AQ1", "AR1", "AS1", "AQ2", "AR2", "AS2", "AQ3", "AR3", "AS3", "AQ4", "AR4", "AS4"].map((address) =>
          document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || "",
        ),
        query,
        rawBasic: window.__webSheetDevCellModel?.("PQ", 2, 44)?.raw,
        rawPro: window.__webSheetDevCellModel?.("PQ", 4, 45)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A57:B60");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Value"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Value");
    await page.locator('[data-data-dialog-action="query-transform-number-absolute"]').click();
    await page.locator("[data-data-field='queryNumberOperand']").fill("10");
    await page.locator('[data-data-dialog-action="query-transform-number-integer-divide"]').click();
    await page.locator("[data-data-field='queryNumberOperand']").fill("4");
    await page.locator('[data-data-dialog-action="query-transform-number-modulo"]').click();
    await page.locator('[data-data-dialog-action="query-transform-number-percentage"]').click();
    await page.locator("[data-data-field='queryNumberOperand']").fill("0.01");
    await page.locator('[data-data-dialog-action="query-transform-number-percent-of"]').click();
    await page.locator("[data-data-field='queryNumberOperand']").fill("2");
    await page.locator('[data-data-dialog-action="query-transform-number-power"]').click();
    const advancedNumberAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("AT1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => Number(window.__webSheetDevCellModel?.("PQ", 2, 47)?.raw) === 4 &&
        Number(window.__webSheetDevCellModel?.("PQ", 4, 47)?.raw) === 9,
      null,
      { timeout: 5000 },
    );
    const advancedNumberLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 57)?.query || null;
      const cellText = (row, col) => {
        const model = window.__webSheetDevCellModel?.("PQ", row, col);
        if (!model) return "";
        return String(model.display ?? model.raw ?? "");
      };
      return {
        values: [
          cellText(1, 46), cellText(1, 47),
          cellText(2, 46), cellText(2, 47),
          cellText(3, 46), cellText(3, 47),
          cellText(4, 46), cellText(4, 47),
        ],
        query,
        rawFirst: window.__webSheetDevCellModel?.("PQ", 2, 47)?.raw,
        rawLast: window.__webSheetDevCellModel?.("PQ", 4, 47)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A63:K65");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "EndYear"));
    const dateTransforms = [
      ["Quarter", "query-transform-date-quarter"],
      ["MonthName", "query-transform-date-month-name"],
      ["DayName", "query-transform-date-day-name"],
      ["DayOfWeek", "query-transform-date-day-of-week"],
      ["DayOfYear", "query-transform-date-day-of-year"],
      ["StartMonth", "query-transform-date-start-month"],
      ["EndMonth", "query-transform-date-end-month"],
      ["StartQuarter", "query-transform-date-start-quarter"],
      ["EndQuarter", "query-transform-date-end-quarter"],
      ["StartYear", "query-transform-date-start-year"],
      ["EndYear", "query-transform-date-end-year"],
    ];
    for (const [column, action] of dateTransforms) {
      await page.locator("[data-data-field='queryActionColumn']").selectOption(column);
      await page.locator(`[data-data-dialog-action="${action}"]`).click();
    }
    const advancedDateAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("AW1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => window.__webSheetDevCellModel?.("PQ", 2, 49)?.raw === 2 &&
        window.__webSheetDevCellModel?.("PQ", 3, 53)?.raw === 365 &&
        window.__webSheetDevCellModel?.("PQ", 2, 59)?.raw === "2026-12-31",
      null,
      { timeout: 5000 },
    );
    const advancedDateLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 63)?.query || null;
      const cellText = (row, col) => {
        const model = window.__webSheetDevCellModel?.("PQ", row, col);
        if (!model) return "";
        return String(model.display ?? model.raw ?? "");
      };
      return {
        headers: Array.from({ length: 11 }, (_item, index) => cellText(1, 49 + index)),
        firstRow: Array.from({ length: 11 }, (_item, index) => cellText(2, 49 + index)),
        secondRow: Array.from({ length: 11 }, (_item, index) => cellText(3, 49 + index)),
        query,
        rawQuarter: window.__webSheetDevCellModel?.("PQ", 2, 49)?.raw,
        rawDayOfYear: window.__webSheetDevCellModel?.("PQ", 3, 53)?.raw,
        rawEndYear: window.__webSheetDevCellModel?.("PQ", 2, 59)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });

    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    await page.locator("[data-data-field='sourceRange']").fill("A68:A71");
    await page.locator('[data-data-dialog-action="transform"]').click();
    await page.waitForSelector('[data-data-dialog="query-transform-editor"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="query-transform-promote-headers"]').click();
    await page.waitForFunction(() => [...document.querySelectorAll("[data-query-transform-column]")].some((input) => input.value === "Text"));
    await page.locator("[data-data-field='queryActionColumn']").selectOption("Text");
    await page.locator("[data-data-field='queryExtractDelimiter']").fill("[");
    await page.locator("[data-data-field='queryExtractEndDelimiter']").fill("]");
    await page.locator('[data-data-dialog-action="query-transform-extract-between"]').click();
    await page.locator("[data-data-field='queryCharacterSet']").fill("abcdefghijklmnopqrstuvwxyz0123456789");
    await page.locator('[data-data-dialog-action="query-transform-keep-characters"]').click();
    await page.locator("[data-data-field='queryPrefixText']").fill("PX-");
    await page.locator('[data-data-dialog-action="query-transform-add-prefix"]').click();
    await page.locator("[data-data-field='querySuffixText']").fill("-OK");
    await page.locator('[data-data-dialog-action="query-transform-add-suffix"]').click();
    await page.locator("[data-data-field='queryExtractStart']").fill("3");
    await page.locator("[data-data-field='queryExtractCount']").fill("4");
    await page.locator('[data-data-dialog-action="query-transform-extract-range"]').click();
    await page.locator("[data-data-field='queryCharacterSet']").fill("bdf");
    await page.locator('[data-data-dialog-action="query-transform-remove-characters"]').click();
    const advancedTextAfterSteps = await page.evaluate(() => ({
      preview: document.querySelector(".query-transform-preview")?.textContent.replace(/\s+/g, "") || "",
      steps: [...document.querySelectorAll(".power-query-step span")].map((item) => item.textContent.trim()),
      columns: [...document.querySelectorAll("[data-query-transform-column]")].map((input) => input.value),
    }));
    await page.locator('[data-data-dialog-action="query-transform-load-to"]').first().click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    await page.locator("[data-data-field='destination']").fill("BI1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => window.__webSheetDevCellModel?.("PQ", 2, 61)?.raw === "a12" &&
        window.__webSheetDevCellModel?.("PQ", 4, 61)?.raw === "e56",
      null,
      { timeout: 5000 },
    );
    const advancedTextLoad = await page.evaluate(() => {
      const query = (window.__webSheetDevQueries?.() || [])
        .find((entry) => entry.query?.sourceType === "tableRange" && entry.query?.sourceRange?.top === 68)?.query || null;
      const cellText = (row, col) => {
        const model = window.__webSheetDevCellModel?.("PQ", row, col);
        if (!model) return "";
        return String(model.display ?? model.raw ?? "");
      };
      return {
        values: [cellText(1, 61), cellText(2, 61), cellText(3, 61), cellText(4, 61)],
        query,
        rawFirst: window.__webSheetDevCellModel?.("PQ", 2, 61)?.raw,
        rawLast: window.__webSheetDevCellModel?.("PQ", 4, 61)?.raw,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    });
    return {
      tableInitial,
      tableAfterSteps,
      tableLoadTo,
      keepDuplicatesAfterSteps,
      keepDuplicatesLoad,
      keepBlankRowsAfterSteps,
      keepBlankRowsLoad,
      webAfterSteps,
      webLoad,
      webNavigatorInitial,
      webNavigatorSelected,
      webNavigatorAfterSteps,
      webNavigatorLoad,
      jsonNavigatorInitial,
      jsonNavigatorAfterSteps,
      jsonNavigatorLoad,
      xmlInitial,
      xmlLoad,
      textAfterSteps,
      textLoad,
      columnAfterSteps,
      columnLoad,
      reshapeAfterSteps,
      reshapeLoad,
      numericAfterSteps,
      numericLoad,
      transposeAfterSteps,
      transposeLoad,
      rowAfterSteps,
      rowLoad,
      keepErrorsAfterSteps,
      keepErrorsLoad,
      pivotAfterSteps,
      pivotLoad,
      advancedNumberAfterSteps,
      advancedNumberLoad,
      advancedDateAfterSteps,
      advancedDateLoad,
      advancedTextAfterSteps,
      advancedTextLoad,
    };
  } finally {
    await page.close();
  }
}



async function exerciseWhatIfGoalSeekDataTable(context) {
  const workbookPath = await createWhatIfWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?whatIf=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="H3"]', { timeout: 10000 });
    await page.locator('[data-ribbon-tab="data"]').click();

    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    const menu = await page.locator("#whatIfMenu").innerText();
    const whatIfMenuKeyboard = await page.evaluate(() => ({
      text: document.querySelector("#whatIfMenu")?.textContent.replace(/\s+/g, " ").trim() || "",
      accessKeys: [...document.querySelectorAll("#whatIfMenu button")].map((button) => button.dataset.accessKey || ""),
      focusedAction: document.activeElement?.dataset?.whatIfMenuAction || "",
    }));
    await page.keyboard.press("G");
    await page.waitForSelector('[data-data-dialog="what-if"]', { timeout: 5000 });
    Object.assign(whatIfMenuKeyboard, await page.evaluate(() => ({
      menuClosed: !document.querySelector("#whatIfMenu"),
      editingCellCount: document.querySelectorAll(".sheet-cell[contenteditable='true']").length,
      dialogTitle: document.querySelector("#whatIfDialogTitle")?.textContent || "",
    })));
    await page.locator("[data-data-field='setCell']").fill("B1");
    await page.locator("[data-data-field='targetValue']").fill("50");
    await page.locator("[data-data-field='changingCell']").fill("B2");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const invalidGoalSeek = await page.evaluate(() => ({
      message: document.querySelector("#messageBoxText")?.textContent || "",
      dialogVisible: Boolean(document.querySelector('[data-data-dialog="what-if"]')),
      b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
    }));
    await page.locator("#messageBoxOverlay [data-message-box-close]").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await page.locator("[data-data-field='setCell']").fill("B2");
    await page.locator("[data-data-field='targetValue']").fill("50");
    await page.locator("[data-data-field='changingCell']").fill("B1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="what-if"]', { state: "detached", timeout: 5000 });
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const goalSeekResult = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("#messageBoxOverlay [data-message-box-close]")].map((button) => button.textContent.trim()),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-close]").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => Number(document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "0") > 24.9,
      null,
      { timeout: 5000 },
    );
    const goalSeek = await page.evaluate(() => ({
      b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
      b2: document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    await page.locator('[data-what-if-menu-action="goal-seek"]').click();
    await page.waitForSelector('[data-data-dialog="what-if"]', { timeout: 5000 });
    await page.locator("[data-data-field='setCell']").fill("WhatIf!C2");
    await page.locator("[data-data-field='targetValue']").fill("60");
    await page.locator("[data-data-field='changingCell']").fill("ExternalInput");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="what-if"]', { state: "detached", timeout: 5000 });
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const crossGoalSeek = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent.trim() || "",
      inputsB1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator("#messageBoxOverlay [data-message-box-close]").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await page.locator(".sheet-tab", { hasText: "WhatIf" }).click();
    await page.waitForSelector('.sheet-cell[data-address="H3"]', { timeout: 5000 });

    await page.locator('.sheet-cell[data-address="G2"]').click();
    await page.locator('.sheet-cell[data-address="H3"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    await page.locator('[data-what-if-menu-action="data-table"]').click();
    await page.waitForSelector('[data-data-dialog="data-table"]', { timeout: 5000 });
    await page.locator("[data-data-field='columnInput']").fill("B1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const invalidDataTable = await page.evaluate(() => ({
      message: document.querySelector("#messageBoxText")?.textContent || "",
      dialogVisible: Boolean(document.querySelector('[data-data-dialog="data-table"]')),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-close]").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await page.locator('[data-data-dialog-action="cancel"]').click();

    await page.locator('.sheet-cell[data-address="D1"]').click();
    await page.locator('.sheet-cell[data-address="E4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    await page.locator('[data-what-if-menu-action="data-table"]').click();
    await page.waitForSelector('[data-data-dialog="data-table"]', { timeout: 5000 });
    await page.locator("[data-data-field='columnInput']").fill("B1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-table"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="E4"] .cell-content')?.textContent === "30",
      null,
      { timeout: 5000 },
    );
    await page.locator('.sheet-cell[data-address="E2"]').click();
    const dataTable = await page.evaluate(() => ({
      values: ["E2", "E3", "E4"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
      formula: document.querySelector("#formulaInput")?.value || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('.sheet-cell[data-address="J1"]').click();
    await page.locator('.sheet-cell[data-address="L3"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    await page.locator('[data-what-if-menu-action="data-table"]').click();
    await page.waitForSelector('[data-data-dialog="data-table"]', { timeout: 5000 });
    await page.locator("[data-data-field='rowInput']").fill("B1");
    await page.locator("[data-data-field='columnInput']").fill("B3");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-table"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="L3"] .cell-content')?.textContent === "12",
      null,
      { timeout: 5000 },
    );
    await page.locator('.sheet-cell[data-address="K2"]').click();
    const twoVariableDataTable = await page.evaluate(() => ({
      values: ["K2", "L2", "K3", "L3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
      b3: document.querySelector('.sheet-cell[data-address="B3"] .cell-content')?.textContent || "",
      formula: document.querySelector("#formulaInput")?.value || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('.sheet-cell[data-address="N1"]').click();
    await page.locator('.sheet-cell[data-address="O4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    await page.locator('[data-what-if-menu-action="data-table"]').click();
    await page.waitForSelector('[data-data-dialog="data-table"]', { timeout: 5000 });
    await page.locator("[data-data-field='columnInput']").fill("ExternalInput");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-table"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="O4"] .cell-content')?.textContent === "18",
      null,
      { timeout: 5000 },
    );
    await page.locator('.sheet-cell[data-address="O2"]').click();
    const crossSheetDataTableValues = await page.evaluate(() => ({
      values: ["O2", "O3", "O4"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      formula: document.querySelector("#formulaInput")?.value || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator(".sheet-tab", { hasText: "Inputs" }).click();
    await page.waitForSelector('.sheet-cell[data-address="B1"]', { timeout: 5000 });
    const crossSheetDataTable = await page.evaluate((base) => ({
      ...base,
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent.trim() || "",
      inputsB1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
    }), crossSheetDataTableValues);

    const groupedWhatIfBaseCount = await page.evaluate(() => document.querySelectorAll(".sheet-tab").length);
    await page.click("#addSheetButton");
    await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 1, groupedWhatIfBaseCount, { timeout: 5000 });
    await page.click("#addSheetButton");
    await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count + 2, groupedWhatIfBaseCount, { timeout: 5000 });
    const groupedWhatIfSheets = await page.evaluate(() =>
      [...document.querySelectorAll(".sheet-tab")]
        .slice(-2)
        .map((button) => ({ name: button.textContent.trim(), index: button.dataset.sheetIndex })),
    );
    await page.evaluate(([first, second]) => {
      const fill = (sheetName, rows) => {
        window.__webSheetDevSetActiveSheet?.(sheetName);
        rows.forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
      };
      fill(first.name, [[1, 1, "Input"], [1, 2, 5], [2, 1, "Output"], [2, 2, "=B1*4"]]);
      fill(second.name, [[1, 1, "Input"], [1, 2, 10], [2, 1, "Output"], [2, 2, "=B1*4"]]);
    }, groupedWhatIfSheets);
    await page.click(`.sheet-tab[data-sheet-index="${groupedWhatIfSheets[0].index}"]`);
    await page.click(`.sheet-tab[data-sheet-index="${groupedWhatIfSheets[1].index}"]`, { modifiers: ["Meta"] });
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    await page.locator('[data-what-if-menu-action="goal-seek"]').click();
    await page.waitForSelector('[data-data-dialog="what-if"]', { timeout: 5000 });
    await page.locator("[data-data-field='setCell']").fill("B2");
    await page.locator("[data-data-field='targetValue']").fill("80");
    await page.locator("[data-data-field='changingCell']").fill("B1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="what-if"]', { state: "detached", timeout: 5000 });
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const groupedGoalSeekMessage = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator("#messageBoxOverlay [data-message-box-close]").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await page.waitForFunction(([first, second]) =>
      Math.abs(Number(window.__webSheetDevCellModel?.(first.name, 1, 2)?.raw) - 20) < 0.01 &&
      Math.abs(Number(window.__webSheetDevCellModel?.(second.name, 1, 2)?.raw) - 20) < 0.01,
      groupedWhatIfSheets,
      { timeout: 5000 },
    );
    const readGroupedGoalSeek = async () => page.evaluate(([first, second]) => ({
      firstB1: window.__webSheetDevCellModel?.(first.name, 1, 2)?.raw ?? "",
      secondB1: window.__webSheetDevCellModel?.(second.name, 1, 2)?.raw ?? "",
      firstFormula: window.__webSheetDevCellModel?.(first.name, 2, 2)?.formula || "",
      secondFormula: window.__webSheetDevCellModel?.(second.name, 2, 2)?.formula || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }), groupedWhatIfSheets);
    const groupedGoalSeek = await readGroupedGoalSeek();
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(([first, second]) =>
      Math.abs(Number(window.__webSheetDevCellModel?.(first.name, 1, 2)?.raw) - 5) < 0.01 &&
      Math.abs(Number(window.__webSheetDevCellModel?.(second.name, 1, 2)?.raw) - 10) < 0.01,
      groupedWhatIfSheets,
      { timeout: 5000 },
    );
    const groupedGoalSeekUndo = await readGroupedGoalSeek();
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(([first, second]) =>
      Math.abs(Number(window.__webSheetDevCellModel?.(first.name, 1, 2)?.raw) - 20) < 0.01 &&
      Math.abs(Number(window.__webSheetDevCellModel?.(second.name, 1, 2)?.raw) - 20) < 0.01,
      groupedWhatIfSheets,
      { timeout: 5000 },
    );
    const groupedGoalSeekRedo = await readGroupedGoalSeek();

    await page.locator(".sheet-tab").first().click();
    await page.evaluate(([first, second]) => {
      const fill = (sheetName, rows) => {
        window.__webSheetDevSetActiveSheet?.(sheetName);
        rows.forEach(([row, col, value]) => window.__webSheetDevSetCell?.(row, col, value));
      };
      fill(first.name, [[1, 4, "Input"], [1, 5, "=B1*2"], [2, 4, 1], [3, 4, 2], [4, 4, 3]]);
      fill(second.name, [[1, 4, "Input"], [1, 5, "=B1*3"], [2, 4, 1], [3, 4, 2], [4, 4, 3]]);
    }, groupedWhatIfSheets);
    await page.click(`.sheet-tab[data-sheet-index="${groupedWhatIfSheets[0].index}"]`);
    await page.click(`.sheet-tab[data-sheet-index="${groupedWhatIfSheets[1].index}"]`, { modifiers: ["Meta"] });
    await page.evaluate(() => window.__webSheetDevSelectRange?.(1, 4, 4, 5));
    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    await page.locator('[data-what-if-menu-action="data-table"]').click();
    await page.waitForSelector('[data-data-dialog="data-table"]', { timeout: 5000 });
    await page.locator("[data-data-field='columnInput']").fill("B1");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="data-table"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(([first, second]) =>
      Number(window.__webSheetDevCellModel?.(first.name, 4, 5)?.raw) === 6 &&
      Number(window.__webSheetDevCellModel?.(second.name, 4, 5)?.raw) === 9,
      groupedWhatIfSheets,
      { timeout: 5000 },
    );
    const readGroupedDataTable = async () => page.evaluate(([first, second]) => {
      const read = (sheetInfo) => ({
        values: [2, 3, 4].map((row) => window.__webSheetDevCellModel?.(sheetInfo.name, row, 5)?.raw ?? ""),
        formulas: [2, 3, 4].map((row) => window.__webSheetDevCellModel?.(sheetInfo.name, row, 5)?.dataTableFormula || ""),
      });
      return {
        first: read(first),
        second: read(second),
        status: document.querySelector("#statusText")?.textContent || "",
      };
    }, groupedWhatIfSheets);
    const groupedDataTable = await readGroupedDataTable();
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(([first, second]) =>
      !window.__webSheetDevCellModel?.(first.name, 2, 5)?.raw &&
      !window.__webSheetDevCellModel?.(second.name, 2, 5)?.raw,
      groupedWhatIfSheets,
      { timeout: 5000 },
    );
    const groupedDataTableUndo = await readGroupedDataTable();
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(([first, second]) =>
      Number(window.__webSheetDevCellModel?.(first.name, 4, 5)?.raw) === 6 &&
      Number(window.__webSheetDevCellModel?.(second.name, 4, 5)?.raw) === 9,
      groupedWhatIfSheets,
      { timeout: 5000 },
    );
    const groupedDataTableRedo = await readGroupedDataTable();
    await page.click(`.sheet-tab[data-sheet-index="${groupedWhatIfSheets[0].index}"]`);
    await page.click(`.sheet-tab[data-sheet-index="${groupedWhatIfSheets[1].index}"]`, { modifiers: ["Meta"] });
    await page.click(`.sheet-tab[data-sheet-index="${groupedWhatIfSheets[0].index}"]`, { button: "right" });
    await page.click('#sheetTabContextMenu [data-sheet-menu-action="delete"]');
    await page.waitForSelector("#messageBox", { timeout: 5000 });
    await page.locator('[data-message-box-result="delete"]').click();
    await page.waitForFunction((count) => document.querySelectorAll(".sheet-tab").length === count, groupedWhatIfBaseCount, { timeout: 5000 });
    const groupedWhatIfCleanup = await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent.trim()),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    return {
      menu,
      whatIfMenuKeyboard,
      invalidGoalSeek,
      goalSeekResult,
      goalSeek,
      crossGoalSeek,
      invalidDataTable,
      dataTable,
      twoVariableDataTable,
      crossSheetDataTable,
      groupedGoalSeekMessage,
      groupedGoalSeek,
      groupedGoalSeekUndo,
      groupedGoalSeekRedo,
      groupedDataTable,
      groupedDataTableUndo,
      groupedDataTableRedo,
      groupedWhatIfCleanup,
    };
  } finally {
    await page.close();
  }
}



async function exerciseScenarioManagerSummary(context) {
  const workbookPath = await createWhatIfWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?scenarioSummary=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B2"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="B1"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    await page.locator('[data-what-if-menu-action="scenario-manager"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-manager"]', { timeout: 5000 });

    await page.locator('[data-data-dialog-action="scenario-add"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-edit"]', { timeout: 5000 });
    await page.locator("[data-data-field='scenarioName']").fill("Low");
    await page.locator("[data-data-field='scenarioCells']").fill("MainInput");
    await page.locator("[data-data-field='scenarioValues']").fill("MainInput=5");
    await page.locator("[data-data-field='scenarioComment']").fill("low case");
    await page.locator('[data-data-dialog-action="scenario-save"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-manager"]', { timeout: 5000 });

    await page.locator('[data-data-dialog-action="scenario-add"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-edit"]', { timeout: 5000 });
    await page.locator("[data-data-field='scenarioName']").fill("Low");
    await page.locator("[data-data-field='scenarioCells']").fill("MainInput");
    await page.locator("[data-data-field='scenarioValues']").fill("MainInput=7");
    await page.locator('[data-data-dialog-action="scenario-save"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const duplicateName = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      dialogVisible: Boolean(document.querySelector('[data-data-dialog="scenario-edit"]')),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-close]").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    await page.locator("[data-data-field='scenarioName']").fill("High");
    await page.locator("[data-data-field='scenarioValues']").fill("MainInput=20");
    await page.locator('[data-data-dialog-action="scenario-save"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-manager"]', { timeout: 5000 });

    await page.locator("[data-data-field='scenarioList']").selectOption({ label: "High" });
    await page.locator('[data-data-dialog-action="scenario-delete"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const deletePrompt = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("#messageBoxOverlay [data-message-box-close]")].map((button) => button.textContent.trim()),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-result='cancel']").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });

    await page.locator("[data-data-field='scenarioList']").selectOption({ label: "High" });
    await page.locator('[data-data-dialog-action="scenario-delete"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    await page.locator("#messageBoxOverlay [data-message-box-result='delete']").click();
    await page.waitForFunction(
      () => [...document.querySelectorAll("[data-data-field='scenarioList'] option")].map((option) => option.textContent.trim()).join("|") === "Low",
      null,
      { timeout: 5000 },
    );
    const deleteUndoBefore = await page.evaluate(() =>
      [...document.querySelectorAll("[data-data-field='scenarioList'] option")].map((option) => option.textContent.trim()),
    );
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => [...document.querySelectorAll("[data-data-field='scenarioList'] option")].map((option) => option.textContent.trim()).join("|") === "Low|High",
      null,
      { timeout: 5000 },
    );
    const deleteUndoAfter = await page.evaluate(() =>
      [...document.querySelectorAll("[data-data-field='scenarioList'] option")].map((option) => option.textContent.trim()),
    );

    await page.locator("[data-data-field='scenarioList']").selectOption({ label: "High" });
    await page.locator('[data-data-dialog-action="scenario-edit"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-edit"]', { timeout: 5000 });
    await page.locator("[data-data-field='scenarioName']").fill("High2");
    await page.locator("[data-data-field='scenarioValues']").fill("MainInput=30");
    await page.locator('[data-data-dialog-action="scenario-save"]').click();
    await page.waitForFunction(
      () => [...document.querySelectorAll("[data-data-field='scenarioList'] option")].map((option) => option.textContent.trim()).join("|") === "Low|High2",
      null,
      { timeout: 5000 },
    );
    const editUndoBefore = await page.evaluate(() =>
      [...document.querySelectorAll("[data-data-field='scenarioList'] option")].map((option) => option.textContent.trim()),
    );
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => [...document.querySelectorAll("[data-data-field='scenarioList'] option")].map((option) => option.textContent.trim()).join("|") === "Low|High",
      null,
      { timeout: 5000 },
    );
    const editUndoAfter = await page.evaluate(() =>
      [...document.querySelectorAll("[data-data-field='scenarioList'] option")].map((option) => option.textContent.trim()),
    );

    const manager = await page.evaluate(() => ({
      options: [...document.querySelectorAll("[data-data-field='scenarioList'] option")].map((option) => option.textContent.trim()),
      description: document.querySelector("[data-scenario-description]")?.textContent || "",
      summaryEnabled: !document.querySelector("[data-data-dialog-action='scenario-summary']")?.disabled,
      b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
      b2: document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent || "",
    }));

    await page.locator("[data-data-field='scenarioList']").selectOption({ label: "High" });
    await page.locator('[data-data-dialog-action="scenario-show"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-manager"]', { state: "detached", timeout: 5000 });
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "20" &&
        document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent === "40",
      null,
      { timeout: 5000 },
    );
    const scenarioShow = await page.evaluate(() => ({
      b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
      b2: document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "10" &&
        document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent === "20",
      null,
      { timeout: 5000 },
    );
    const scenarioShowUndo = await page.evaluate(() => ({
      b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
      b2: document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "20" &&
        document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent === "40",
      null,
      { timeout: 5000 },
    );
    const scenarioShowRedo = await page.evaluate(() => ({
      b1: document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent || "",
      b2: document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B1"] .cell-content')?.textContent === "10" &&
        document.querySelector('.sheet-cell[data-address="B2"] .cell-content')?.textContent === "20",
      null,
      { timeout: 5000 },
    );

    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    await page.locator('[data-what-if-menu-action="scenario-manager"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-manager"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="scenario-summary"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-summary"]', { timeout: 5000 });
    const summaryDialog = await page.evaluate(() => ({
      title: document.querySelector("#scenarioSummaryTitle")?.textContent || "",
      reportTypes: [...document.querySelectorAll("[data-data-field='reportType']")].map((input) => ({
        value: input.value,
        checked: input.checked,
        label: input.closest("label")?.textContent?.trim() || "",
      })),
      defaultResultCells: document.querySelector("[data-data-field='resultCells']")?.value || "",
    }));
    await page.locator("[data-data-field='resultCells']").fill("MainResult");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim()?.startsWith("シナリオの概要"),
      null,
      { timeout: 5000 },
    );
    const summary = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      title: document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent || "",
      headers: ["A3", "B3", "C3", "D3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      changing: ["A4", "B4", "C4", "D4"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      resultHeader: ["A6", "B6", "C6", "D6"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      result: ["A7", "B7", "C7", "D7"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "WhatIf",
      null,
      { timeout: 5000 },
    );
    const summaryUndo = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      sheetNames: [...document.querySelectorAll(".sheet-tab")].map((button) => button.textContent.trim()),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim()?.startsWith("シナリオの概要"),
      null,
      { timeout: 5000 },
    );
    const summaryRedo = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      title: document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator(".sheet-tab").filter({ hasText: "WhatIf" }).click();
    await page.locator('[data-ribbon-command="data-what-if"]').click();
    await page.waitForSelector("#whatIfMenu", { timeout: 5000 });
    await page.locator('[data-what-if-menu-action="scenario-manager"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-manager"]', { timeout: 5000 });
    await page.locator('[data-data-dialog-action="scenario-summary"]').click();
    await page.waitForSelector('[data-data-dialog="scenario-summary"]', { timeout: 5000 });
    await page.locator("[data-data-field='reportType'][value='pivot']").check();
    await page.locator("[data-data-field='resultCells']").fill("MainResult");
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForFunction(
      () => document.querySelector(".sheet-tab.is-active")?.textContent?.trim()?.startsWith("シナリオ ピボット"),
      null,
      { timeout: 5000 },
    );
    const pivotSummary = await page.evaluate(() => ({
      activeSheet: document.querySelector(".sheet-tab.is-active")?.textContent?.trim() || "",
      title: document.querySelector('.sheet-cell[data-address="A1"] .cell-content')?.textContent || "",
      groups: ["A2", "B2", "C2"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      headers: ["A3", "B3", "C3"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      current: ["A4", "B4", "C4"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      low: ["A5", "B5", "C5"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      high: ["A6", "B6", "C6"].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    return {
      duplicateName,
      deletePrompt,
      deleteUndoBefore,
      deleteUndoAfter,
      editUndoBefore,
      editUndoAfter,
      manager,
      scenarioShow,
      scenarioShowUndo,
      scenarioShowRedo,
      summaryDialog,
      summary,
      summaryUndo,
      summaryRedo,
      pivotSummary,
    };
  } finally {
    await page.close();
  }
}



async function exerciseFlashFillDataPatterns(context) {
  const workbookPath = await createFlashFillWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?flashFill=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="R4"]', { timeout: 10000 });
    await page.locator('[data-ribbon-tab="data"]').click();
    const flashFillRange = async (start, end, expectedAddress, expectedText) => {
      await page.locator(`.sheet-cell[data-address="${start}"]`).click();
      await page.locator(`.sheet-cell[data-address="${end}"]`).click({ modifiers: ["Shift"] });
      await page.locator('[data-ribbon-command="data-flash-fill"]').click();
      await page.waitForFunction(
        ({ address, text }) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent === text,
        { address: expectedAddress, text: expectedText },
        { timeout: 5000 },
      );
    };
    await page.locator('.sheet-cell[data-address="B2"]').click();
    await page.locator('.sheet-cell[data-address="B4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-flash-fill"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="B3"] .cell-content')?.textContent === "AL",
      null,
      { timeout: 5000 },
    );
    await page.locator('.sheet-cell[data-address="D2"]').click();
    await page.locator('.sheet-cell[data-address="D4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-flash-fill"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="D3"] .cell-content')?.textContent === "beta.long",
      null,
      { timeout: 5000 },
    );
    await page.locator('.sheet-cell[data-address="F2"]').click();
    await page.locator('.sheet-cell[data-address="F4"]').click({ modifiers: ["Shift"] });
    await page.locator('[data-ribbon-command="data-flash-fill"]').click();
    await page.waitForFunction(
      () => document.querySelector('.sheet-cell[data-address="F3"] .cell-content')?.textContent === "Lovelace, Ada",
      null,
      { timeout: 5000 },
    );
    await flashFillRange("H2", "H4", "H3", "120");
    await flashFillRange("J2", "J4", "J3", "CD");
    await flashFillRange("L2", "L4", "L3", "OS");
    await flashFillRange("N2", "N4", "N3", "2026-06-01");
    await flashFillRange("P2", "P4", "P3", "080-9988-7766");
    await flashFillRange("R2", "R4", "R3", "2026/06/01");
    const values = await page.evaluate(() =>
      [
        "B2", "B3", "B4",
        "D2", "D3", "D4",
        "F2", "F3", "F4",
        "H2", "H3", "H4",
        "J2", "J3", "J4",
        "L2", "L3", "L4",
        "N2", "N3", "N4",
        "P2", "P3", "P4",
        "R2", "R3", "R4",
      ].map((address) => document.querySelector(`.sheet-cell[data-address="${address}"] .cell-content`)?.textContent || ""),
    );
    const status = await page.locator("#statusText").innerText();
    return { values, status };
  } finally {
    await page.close();
  }
}



async function exerciseDataCurrentRegion(context) {
  const workbookPath = await createAutoFilterMenuWorkbook();
  const filterPage = await context.newPage();
  const rowHeaderFilterPage = await context.newPage();
  const chtHeaderFilterPage = await context.newPage();
  const blankHeaderFilterPage = await context.newPage();
  const tablePriorityFilterPage = await context.newPage();
  const toggleFilterPage = await context.newPage();
  const sortPage = await context.newPage();
  const duplicatePage = await context.newPage();
  try {
    await filterPage.goto(`${appUrl}?dataRegionFilter=${Date.now()}`, { waitUntil: "load" });
    await filterPage.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(filterPage, workbookPath);
    await filterPage.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await filterPage.locator('.sheet-cell[data-address="B3"]').click();
    await filterPage.locator('[data-ribbon-tab="data"]').click();
    await filterPage.locator('[data-ribbon-command="data-filter"]').click();
    await filterPage.waitForSelector('.cell-filter-button[data-filter-col="2"]', { timeout: 5000 });
    const filterCols = await filterPage.locator(".cell-filter-button").evaluateAll((buttons) =>
      buttons.map((button) => Number(button.getAttribute("data-filter-col"))).sort((left, right) => left - right),
    );
    const filterStatus = await filterPage.locator("#statusText").innerText();

    await rowHeaderFilterPage.goto(`${appUrl}?dataRegionRowHeaderFilter=${Date.now()}`, { waitUntil: "load" });
    await rowHeaderFilterPage.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(rowHeaderFilterPage, workbookPath);
    await rowHeaderFilterPage.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await rowHeaderFilterPage.locator(".row-header[data-row='1']").click();
    await rowHeaderFilterPage.locator('[data-ribbon-tab="data"]').click();
    await rowHeaderFilterPage.locator('[data-ribbon-command="data-filter"]').click();
    await rowHeaderFilterPage.waitForSelector('.cell-filter-button[data-filter-col="2"]', { timeout: 5000 });
    const rowHeaderFilter = await rowHeaderFilterPage.evaluate(() => ({
      cols: [...document.querySelectorAll(".cell-filter-button")]
        .map((button) => Number(button.getAttribute("data-filter-col")))
        .sort((left, right) => left - right),
      rows: [...document.querySelectorAll(".cell-filter-button")]
        .map((button) => Number(button.getAttribute("data-filter-row")))
        .sort((left, right) => left - right),
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await chtHeaderFilterPage.goto(`${appUrl}?chtHeaderRowFilter=${Date.now()}`, { waitUntil: "load" });
    await chtHeaderFilterPage.setInputFiles("#fileInput", defaultWorkbook);
    await waitForUploadedWorkbook(chtHeaderFilterPage, defaultWorkbook);
    await chtHeaderFilterPage.waitForSelector(".sheet-tab.is-active", { timeout: 10000 });
    const chtTab = chtHeaderFilterPage.locator(".sheet-tab", { hasText: "CHT" }).first();
    if (await chtTab.count()) await chtTab.click();
    await chtHeaderFilterPage.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "CHT", null, { timeout: 10000 });
    await chtHeaderFilterPage.evaluate(() => {
      const host = document.querySelector("#sheetHost");
      if (!host) return;
      host.scrollTop = 0;
      host.scrollLeft = 0;
      host.dispatchEvent(new Event("scroll"));
    });
    await chtHeaderFilterPage.waitForSelector('.row-header[data-row="15"]', { timeout: 10000 });
    await chtHeaderFilterPage.locator('.row-header[data-row="15"]').click();
    await chtHeaderFilterPage.locator('[data-ribbon-tab="data"]').click();
    await chtHeaderFilterPage.locator('[data-ribbon-command="data-filter"]').click();
    await chtHeaderFilterPage.waitForSelector('.cell-filter-button[data-filter-row="15"]', { timeout: 5000 });
    const chtHeaderFilter = await chtHeaderFilterPage.evaluate(() => ({
      rows: [...document.querySelectorAll(".cell-filter-button")]
        .map((button) => Number(button.getAttribute("data-filter-row")))
        .sort((left, right) => left - right),
      cols: [...document.querySelectorAll(".cell-filter-button")]
        .map((button) => Number(button.getAttribute("data-filter-col")))
        .sort((left, right) => left - right),
      modelRange: window.__webSheetDevSheetModel?.("CHT")?.autoFilter?.range || null,
      label: document.querySelector("#selectedCellLabel")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await blankHeaderFilterPage.goto(`${appUrl}?dataRegionBlankHeaderFilter=${Date.now()}`, { waitUntil: "load" });
    await blankHeaderFilterPage.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(blankHeaderFilterPage, workbookPath);
    await blankHeaderFilterPage.evaluate(() => window.__webSheetDevSetActiveSheet?.("BlankHeader"));
    await blankHeaderFilterPage.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "BlankHeader", null, { timeout: 10000 });
    await blankHeaderFilterPage.evaluate(() => window.__webSheetDevSelectRange?.(1, 1, 1, 3));
    await blankHeaderFilterPage.locator('[data-ribbon-tab="data"]').click();
    await blankHeaderFilterPage.locator('[data-ribbon-command="data-filter"]').click();
    await blankHeaderFilterPage.waitForSelector('.cell-filter-button[data-filter-col="2"]', { timeout: 5000 });
    const blankHeaderFilter = await blankHeaderFilterPage.evaluate(() => ({
      cols: [...document.querySelectorAll(".cell-filter-button")]
        .map((button) => Number(button.getAttribute("data-filter-col")))
        .sort((left, right) => left - right),
      rows: [...document.querySelectorAll(".cell-filter-button")]
        .map((button) => Number(button.getAttribute("data-filter-row")))
        .sort((left, right) => left - right),
      range: window.__webSheetDevSheetModel?.("BlankHeader")?.autoFilter?.range || null,
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await tablePriorityFilterPage.goto(`${appUrl}?dataRegionTablePriority=${Date.now()}`, { waitUntil: "load" });
    await tablePriorityFilterPage.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(tablePriorityFilterPage, workbookPath);
    await tablePriorityFilterPage.evaluate(() => window.__webSheetDevSetActiveSheet?.("TablePriority"));
    await tablePriorityFilterPage.waitForFunction(() => document.querySelector(".sheet-tab.is-active")?.textContent?.trim() === "TablePriority", null, { timeout: 10000 });
    await tablePriorityFilterPage.evaluate(() => window.__webSheetDevSelectRange?.(3, 3, 3, 3));
    await tablePriorityFilterPage.locator('[data-ribbon-tab="data"]').click();
    const tablePriorityInitiallyFiltered = await tablePriorityFilterPage.evaluate(() => Boolean(window.__webSheetDevSheetModel?.("TablePriority")?.autoFilter?.range));
    await tablePriorityFilterPage.locator('[data-ribbon-command="data-filter"]').click();
    if (tablePriorityInitiallyFiltered) {
      await tablePriorityFilterPage.waitForFunction(() => !window.__webSheetDevSheetModel?.("TablePriority")?.autoFilter?.range, null, { timeout: 5000 });
      await tablePriorityFilterPage.locator('[data-ribbon-command="data-filter"]').click();
    }
    await tablePriorityFilterPage.waitForSelector('.cell-filter-button[data-filter-col="4"]', { timeout: 5000 });
    const tablePriorityFilter = await tablePriorityFilterPage.evaluate(() => ({
      cols: [...document.querySelectorAll(".cell-filter-button")]
        .map((button) => Number(button.getAttribute("data-filter-col")))
        .sort((left, right) => left - right),
      range: window.__webSheetDevSheetModel?.("TablePriority")?.autoFilter?.range || null,
      table: window.__webSheetDevSheetModel?.("TablePriority")?.tables?.[0] || null,
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await toggleFilterPage.goto(`${appUrl}?dataRegionFilterToggle=${Date.now()}`, { waitUntil: "load" });
    await toggleFilterPage.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(toggleFilterPage, workbookPath);
    await toggleFilterPage.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await toggleFilterPage.evaluate(() => window.__webSheetDevSelectRange?.(3, 2, 3, 2));
    await toggleFilterPage.locator('[data-ribbon-tab="data"]').click();
    await toggleFilterPage.locator('[data-ribbon-command="data-filter"]').click();
    await toggleFilterPage.waitForSelector('.cell-filter-button[data-filter-col="2"]', { timeout: 5000 });
    await toggleFilterPage.evaluate(() => window.__webSheetDevSelectRange?.(3, 2, 3, 2));
    await toggleFilterPage.locator('[data-ribbon-command="data-filter"]').click();
    await toggleFilterPage.waitForFunction(() => !window.__webSheetDevSheetModel?.("Filter")?.autoFilter?.range, null, { timeout: 5000 });
    const toggleFilter = await toggleFilterPage.evaluate(() => ({
      hasAutoFilter: Boolean(window.__webSheetDevSheetModel?.("Filter")?.autoFilter?.range),
      buttonCount: document.querySelectorAll(".cell-filter-button").length,
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await sortPage.goto(`${appUrl}?dataRegionSort=${Date.now()}`, { waitUntil: "load" });
    await sortPage.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(sortPage, workbookPath);
    await sortPage.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await sortPage.locator('.sheet-cell[data-address="B3"]').click();
    await sortPage.locator('[data-ribbon-tab="data"]').click();
    await sortPage.locator('[data-ribbon-command="data-sort-dialog"]').click();
    await sortPage.waitForSelector('[data-data-dialog="sort"]', { timeout: 5000 });
    const sortKeys = await sortPage.locator('[data-data-sort-key="0"] option').evaluateAll((options) =>
      options.map((option) => option.textContent.trim()),
    );

    await duplicatePage.goto(`${appUrl}?dataRegionDuplicates=${Date.now()}`, { waitUntil: "load" });
    await duplicatePage.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(duplicatePage, workbookPath);
    await duplicatePage.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await duplicatePage.locator('.sheet-cell[data-address="B3"]').click();
    await duplicatePage.locator('[data-ribbon-tab="data"]').click();
    await duplicatePage.locator('[data-ribbon-command="data-remove-duplicates"]').click();
    await duplicatePage.waitForSelector('[data-data-dialog="remove-duplicates"]', { timeout: 5000 });
    const duplicateLabels = await duplicatePage.locator("[data-duplicate-col-label]").evaluateAll((labels) =>
      labels.map((label) => label.textContent.trim()),
    );
    return { filterCols, filterStatus, rowHeaderFilter, chtHeaderFilter, blankHeaderFilter, tablePriorityFilter, toggleFilter, sortKeys, duplicateLabels };
  } finally {
    await filterPage.close();
    await rowHeaderFilterPage.close();
    await chtHeaderFilterPage.close();
    await blankHeaderFilterPage.close();
    await tablePriorityFilterPage.close();
    await toggleFilterPage.close();
    await sortPage.close();
    await duplicatePage.close();
  }
}



async function exerciseQueryConnectionsPane(context) {
  const workbookPath = await createAutoFilterMenuWorkbook();
  const page = await context.newPage();
  try {
    await page.goto(`${appUrl}?queryConnections=${Date.now()}`, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await waitForUploadedWorkbook(page, workbookPath);
    await page.waitForSelector('.sheet-cell[data-address="B5"]', { timeout: 10000 });
    await page.locator('.sheet-cell[data-address="B3"]').click();
    await page.locator('[data-ribbon-tab="data"]').click();
    await page.locator('[data-ribbon-command="data-get-table-range"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { timeout: 5000 });
    const loadDialog = await page.evaluate(() => ({
      range: document.querySelector("[data-data-field='sourceRange']")?.value || "",
      hasHeader: document.querySelector("[data-data-field='hasHeader']")?.checked || false,
      previewText: document.querySelector(".query-preview")?.textContent?.replace(/\s+/g, " ").trim() || "",
    }));
    await page.locator('[data-data-dialog-action="load"]').click();
    await page.waitForSelector('[data-data-dialog="get-table-range"]', { state: "detached", timeout: 5000 });
    await page.locator('[data-ribbon-command="data-queries-connections"]').click();
    await page.waitForSelector("#queryPane", { timeout: 5000 });
    const initialPane = await page.evaluate(() => ({
      title: document.querySelector(".query-pane-header strong")?.textContent || "",
      tabs: [...document.querySelectorAll(".query-pane-tabs button")].map((button) => ({
        text: button.textContent.replace(/\s+/g, " ").trim(),
        active: button.classList.contains("is-active"),
      })),
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      status: document.querySelector("#statusText")?.textContent || "",
    }));

    await page.locator('[data-query-pane-action="load-to"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { timeout: 5000 });
    const loadToDialog = await page.evaluate(() => ({
      title: document.querySelector("#importDataTitle")?.textContent || "",
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
      propertyButton: document.querySelector("[data-data-dialog-action='import-properties']")?.textContent.trim() || "",
    }));
    await page.locator('[data-data-dialog-action="cancel"]').click();
    await page.waitForSelector('[data-data-dialog="import-data"]', { state: "detached", timeout: 5000 });

    await page.locator('[data-query-pane-action="properties"]').click();
    await page.waitForSelector('[data-data-dialog="query-properties"]', { timeout: 5000 });
    const propertiesInitial = await page.evaluate(() => ({
      tabs: [...document.querySelectorAll("[data-query-properties-tab]")].map((button) => ({
        text: button.textContent.trim(),
        active: button.classList.contains("is-active"),
      })),
      usageLabels: [...document.querySelectorAll("[data-query-properties-tab-panel='usage'] label")].map((label) =>
        label.textContent.replace(/\s+/g, " ").trim(),
      ),
      interval: document.querySelector("[data-data-field='refreshIntervalMinutes']")?.value || "",
      intervalDisabled: Boolean(document.querySelector("[data-data-field='refreshIntervalMinutes']")?.disabled),
      typeNote: document.querySelector("[data-query-properties-tab-panel='usage'] .data-dialog-note")?.textContent || "",
    }));
    await page.locator('[data-query-properties-tab="definition"]').click();
    await page.waitForFunction(
      () => document.querySelector("[data-query-properties-tab-panel='definition']")?.classList.contains("is-active"),
      null,
      { timeout: 5000 },
    );
    const definitionInitial = await page.evaluate(() => ({
      labels: [...document.querySelectorAll("[data-query-properties-tab-panel='definition'] .excel-dialog-caption")].map((label) => label.textContent.trim()),
      connectionString: document.querySelector("[data-data-field='connectionString']")?.value || "",
      commandText: document.querySelector("[data-data-field='commandText']")?.value || "",
      commandType: document.querySelector("[data-data-field='commandType']")?.value || "",
      commandTypeOptions: [...document.querySelectorAll("[data-data-field='commandType'] option")].map((option) => option.textContent.trim()),
      definitionChecks: [...document.querySelectorAll("[data-query-properties-tab-panel='definition'] .data-dialog-check")].map((label) =>
        label.textContent.replace(/\s+/g, " ").trim(),
      ),
    }));
    await page.locator("[data-data-field='name']").fill("SalesQuery");
    await page.locator("[data-data-field='description']").fill("range query");
    await page.locator("[data-data-field='connectionFile']").fill("BookConnection.odc");
    await page.locator("[data-data-field='alwaysUseConnectionFile']").check();
    await page.locator("[data-data-field='failIfConnectionFileMissing']").check();
    await page.locator("[data-data-field='commandType']").selectOption("table");
    await page.locator("[data-data-field='commandText']").fill("Filter!A1:B5");
    await page.locator('[data-query-properties-tab="usage"]').click();
    await page.waitForFunction(
      () => document.querySelector("[data-query-properties-tab-panel='usage']")?.classList.contains("is-active"),
      null,
      { timeout: 5000 },
    );
    await page.locator("[data-data-field='backgroundRefresh']").check();
    await page.locator("[data-data-field='refreshEveryEnabled']").check();
    const intervalEnabledAfterCheck = await page.evaluate(() => !document.querySelector("[data-data-field='refreshIntervalMinutes']")?.disabled);
    await page.locator("[data-data-field='refreshIntervalMinutes']").fill("15");
    await page.locator("[data-data-field='refreshOnOpen']").check();
    await page.locator("[data-data-field='removeDataBeforeSave']").check();
    await page.locator("[data-data-field='refreshWithAll']").uncheck();
    await page.locator('[data-data-dialog-action="ok"]').click();
    await page.waitForSelector('[data-data-dialog="query-properties"]', { state: "detached", timeout: 5000 });
    const renamed = await page.evaluate(() => ({
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      status: document.querySelector("#statusText")?.textContent || "",
      query: window.__webSheetDevQueries?.()[0]?.query || null,
      refreshTimerCount: window.__webSheetQueryRefreshTimerCount?.() ?? -1,
    }));
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => [...document.querySelectorAll(".query-pane-item")].some((row) => row.textContent.includes("Query1")),
      null,
      { timeout: 5000 },
    );
    const propertyUndo = await page.evaluate(() => ({
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      status: document.querySelector("#statusText")?.textContent || "",
      refreshTimerCount: window.__webSheetQueryRefreshTimerCount?.() ?? -1,
    }));
    await page.keyboard.press("Control+Y");
    await page.waitForFunction(
      () => [...document.querySelectorAll(".query-pane-item")].some((row) => row.textContent.includes("SalesQuery")),
      null,
      { timeout: 5000 },
    );
    const propertyRedo = await page.evaluate(() => ({
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      status: document.querySelector("#statusText")?.textContent || "",
      refreshTimerCount: window.__webSheetQueryRefreshTimerCount?.() ?? -1,
    }));

    await page.locator('[data-query-pane-action="refresh-all"]').click();
    await page.waitForFunction(
      () => document.querySelector("#statusText")?.textContent?.includes("更新する接続はありません"),
      null,
      { timeout: 5000 },
    );
    const refreshAllSkipped = await page.locator("#statusText").innerText();

    await page.locator('[data-query-pane-action="tab"][data-query-pane-tab="connections"]').click();
    await page.waitForFunction(
      () => document.querySelector(".query-pane-tabs button.is-active")?.textContent?.includes("接続"),
      null,
      { timeout: 5000 },
    );
    const connectionsTab = await page.evaluate(() => ({
      tabs: [...document.querySelectorAll(".query-pane-tabs button")].map((button) => ({
        text: button.textContent.replace(/\s+/g, " ").trim(),
        active: button.classList.contains("is-active"),
      })),
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
    }));

    await page.locator('[data-query-pane-action="refresh"]').click();
    await page.waitForFunction(
      () => document.querySelector("#statusText")?.textContent?.includes("SalesQuery を更新"),
      null,
      { timeout: 5000 },
    );
    const singleRefresh = await page.evaluate(() => ({
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.locator('[data-query-pane-action="delete"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const connectionDeletePrompt = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("#messageBoxOverlay [data-message-box-close]")].map((button) => button.textContent.trim()),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-result='cancel']").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    const connectionDeleteCancelCount = await page.locator(".query-pane-item").count();

    await page.locator('[data-query-pane-action="tab"][data-query-pane-tab="queries"]').click();
    await page.waitForFunction(
      () => document.querySelector(".query-pane-tabs button.is-active")?.textContent?.includes("クエリ"),
      null,
      { timeout: 5000 },
    );
    await page.locator('[data-query-pane-action="delete"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    const deletePrompt = await page.evaluate(() => ({
      title: document.querySelector("#messageBoxTitle")?.textContent || "",
      message: document.querySelector("#messageBoxText")?.textContent || "",
      buttons: [...document.querySelectorAll("#messageBoxOverlay [data-message-box-close]")].map((button) => button.textContent.trim()),
    }));
    await page.locator("#messageBoxOverlay [data-message-box-result='cancel']").click();
    await page.waitForSelector("#messageBoxOverlay", { state: "detached", timeout: 5000 });
    const afterCancelDeleteCount = await page.locator(".query-pane-item").count();
    await page.locator('[data-query-pane-action="delete"]').click();
    await page.waitForSelector("#messageBoxOverlay", { timeout: 5000 });
    await page.locator("#messageBoxOverlay [data-message-box-result='delete']").click();
    await page.waitForFunction(
      () => document.querySelector(".query-pane-empty")?.textContent?.includes("クエリがありません"),
      null,
      { timeout: 5000 },
    );
    const afterDelete = await page.evaluate(() => ({
      itemCount: document.querySelectorAll(".query-pane-item").length,
      empty: document.querySelector(".query-pane-empty")?.textContent || "",
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    await page.keyboard.press("Control+Z");
    await page.waitForFunction(
      () => [...document.querySelectorAll(".query-pane-item")].some((row) => row.textContent.includes("SalesQuery")),
      null,
      { timeout: 5000 },
    );
    const afterDeleteUndo = await page.evaluate(() => ({
      itemCount: document.querySelectorAll(".query-pane-item").length,
      rows: [...document.querySelectorAll(".query-pane-item")].map((row) => row.textContent.replace(/\s+/g, " ").trim()),
      status: document.querySelector("#statusText")?.textContent || "",
    }));
    return {
      loadDialog,
      initialPane,
      loadToDialog,
      propertiesInitial,
      definitionInitial,
      intervalEnabledAfterCheck,
      renamed,
      propertyUndo,
      propertyRedo,
      refreshAllSkipped,
      connectionsTab,
      singleRefresh,
      connectionDeletePrompt,
      connectionDeleteCancelCount,
      deletePrompt,
      afterCancelDeleteCount,
      afterDelete,
      afterDeleteUndo,
    };
  } finally {
    await page.close();
  }
}

export {
  exerciseDataCurrentRegion,
  exerciseFlashFillDataPatterns,
  exercisePowerQueryTransformEditor,
  exerciseQueryConnectionsPane,
  exerciseScenarioManagerSummary,
  exerciseWhatIfGoalSeekDataTable,
};
