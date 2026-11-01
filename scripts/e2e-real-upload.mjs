// Sequential E2E feature module.
import {
  createCustomViewWorkbook,
  createPageBackgroundImage,
  createRichNativeShapeWorkbook,
} from "./e2e/helpers/fixtures-and-model.mjs";
import {
  runFormulaSmokeOnly,
  runViewSmokeOnly,
} from "./e2e/helpers/formula-smoke.mjs";
import {
  runFormulaUiSmokeOnly,
} from "./e2e/helpers/formula-ui-smoke.mjs";
import {
  runAltEnterSmokeOnly,
  runAutoFilterDateOnly,
  runClipboardSessionSmokeOnly,
  runCustomViewImportExportOnly,
  runDataFilterSortOnly,
  runDatePersistenceSmokeOnly,
  runMacroVbaSubsetOnly,
  runOutlineGroupingOnly,
  runSaveSmokeOnly,
  runValidationMenuOnly,
} from "./e2e/helpers/other-smoke.mjs";
import {
  runSetupAndWorkbook,
} from "./e2e/scenarios/setup-and-workbook.mjs";
import {
  runEditingAndFormat,
} from "./e2e/scenarios/editing-and-format.mjs";
import {
  runWindowsAndSelection,
} from "./e2e/scenarios/windows-and-selection.mjs";
import {
  runObjectsAndFormulas,
} from "./e2e/scenarios/objects-and-formulas.mjs";
import {
  runImportAndRoundtrip,
} from "./e2e/scenarios/import-and-roundtrip.mjs";
import {
  runExportAndResults,
} from "./e2e/scenarios/export-and-results.mjs";
import {
  runAssertionsEditor,
} from "./e2e/scenarios/assertions-editor.mjs";
import {
  runAssertionsWorkbook,
} from "./e2e/scenarios/assertions-workbook.mjs";
import {
  runAssertionsFormulasAndClipboard,
} from "./e2e/scenarios/assertions-formulas-and-clipboard.mjs";
import { chromium } from "@playwright/test";


import ExcelJS from "exceljs";


import { access, mkdtemp, readFile, writeFile } from "node:fs/promises";


import os from "node:os";


import path from "node:path";


import { pathToFileURL } from "node:url";


import JSZip from "jszip";


import { standaloneRuntimeCurrentBaseUrl, standaloneRuntimeCurrentDir, STANDALONE_RUNTIME_CHANNEL } from "./standalone-runtime-path.mjs";


import { gunzipSync } from "node:zlib";



const defaultWorkbook = path.resolve("testfile/ZJ_スキルシート_20260517.xlsx");


const cliArgs = process.argv.slice(2);


const formulaSmokeOnly = cliArgs.includes("--formula-smoke");


const formulaUiSmokeOnly = cliArgs.includes("--formula-ui-smoke");


const viewSmokeOnly = cliArgs.includes("--view-smoke");


const viewDiagnosticsOnly = cliArgs.includes("--view-diagnostics");


const validationMenuOnly = cliArgs.includes("--validation-menu");


const outlineGroupingOnly = cliArgs.includes("--outline-grouping");


const macroVbaSubsetOnly = cliArgs.includes("--macro-vba-subset");


const customViewImportExportOnly = cliArgs.includes("--custom-view-import-export");


const dataFilterSortOnly = cliArgs.includes("--data-filter-sort");


const autoFilterDateOnly = cliArgs.includes("--auto-filter-date");


const saveSmokeOnly = cliArgs.includes("--save-smoke");


const datePersistenceSmokeOnly = cliArgs.includes("--date-persistence-smoke");


const clipboardSessionSmokeOnly = cliArgs.includes("--clipboard-session-smoke");


const altEnterSmokeOnly = cliArgs.includes("--alt-enter-smoke");


const smokeOnly =
  formulaSmokeOnly ||
  formulaUiSmokeOnly ||
  viewSmokeOnly ||
  validationMenuOnly ||
  outlineGroupingOnly ||
  macroVbaSubsetOnly ||
  dataFilterSortOnly ||
  autoFilterDateOnly ||
  saveSmokeOnly ||
  datePersistenceSmokeOnly ||
  clipboardSessionSmokeOnly ||
  altEnterSmokeOnly;


const workbookArgument = cliArgs.find((arg) => arg && !arg.startsWith("--"));


const workbookPath = workbookArgument ? path.resolve(workbookArgument) : defaultWorkbook;


const nativeShapeWorkbookPath = path.resolve("testfile/ZJ_スキルシート_20260517_test用.xlsx");


const appUrl = process.env.WEBSHEET_URL || "http://127.0.0.1:5173/";


const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";


const e2eProgressEnabled = process.env.WEBSHEET_E2E_PROGRESS === "1";



await access(workbookPath);


if (!smokeOnly) await access(nativeShapeWorkbookPath);


const richNativeShapeWorkbookPath = smokeOnly ? nativeShapeWorkbookPath : await createRichNativeShapeWorkbook(nativeShapeWorkbookPath);


const customViewWorkbookPath = smokeOnly ? "" : await createCustomViewWorkbook();


const pageBackgroundImagePath = smokeOnly ? "" : await createPageBackgroundImage();



const WEBSHEET_MODEL_SCRIPT_RE = /<script\b(?=[^>]*\bid="websheet-model")(?=[^>]*\btype="application\/json")[^>]*>([\s\S]*?)<\/script>/;



if (formulaSmokeOnly) {
  await runFormulaSmokeOnly();
  process.exit(0);
}



if (formulaUiSmokeOnly) {
  await runFormulaUiSmokeOnly();
  process.exit(0);
}



if (viewSmokeOnly) {
  await runViewSmokeOnly();
  process.exit(0);
}



if (datePersistenceSmokeOnly) {
  await runDatePersistenceSmokeOnly();
  process.exit(0);
}



if (clipboardSessionSmokeOnly) {
  await runClipboardSessionSmokeOnly();
  process.exit(0);
}



if (altEnterSmokeOnly) {
  await runAltEnterSmokeOnly();
  process.exit(0);
}



if (macroVbaSubsetOnly) {
  await runMacroVbaSubsetOnly();
  process.exit(0);
}



if (validationMenuOnly) {
  await runValidationMenuOnly();
  process.exit(0);
}



if (outlineGroupingOnly) {
  await runOutlineGroupingOnly();
  process.exit(0);
}



if (autoFilterDateOnly) {
  await runAutoFilterDateOnly();
  process.exit(0);
}



if (dataFilterSortOnly) {
  await runDataFilterSortOnly();
  process.exit(0);
}



if (saveSmokeOnly) {
  await runSaveSmokeOnly();
  process.exit(0);
}



if (customViewImportExportOnly) {
  await runCustomViewImportExportOnly();
  process.exit(0);
}



const browser = await chromium.launch({
  headless: true,
  executablePath: chromePath,
});

try {
const scenarioState = {};
await runSetupAndWorkbook(scenarioState);
await runEditingAndFormat(scenarioState);
await runWindowsAndSelection(scenarioState);
await runObjectsAndFormulas(scenarioState);
await runImportAndRoundtrip(scenarioState);
await runExportAndResults(scenarioState);
await runAssertionsEditor(scenarioState);
await runAssertionsWorkbook(scenarioState);
await runAssertionsFormulasAndClipboard(scenarioState);
} finally {
  await browser.close();
}

export {
  ExcelJS,
  JSZip,
  STANDALONE_RUNTIME_CHANNEL,
  WEBSHEET_MODEL_SCRIPT_RE,
  access,
  altEnterSmokeOnly,
  appUrl,
  autoFilterDateOnly,
  browser,
  chromePath,
  chromium,
  cliArgs,
  clipboardSessionSmokeOnly,
  customViewImportExportOnly,
  customViewWorkbookPath,
  dataFilterSortOnly,
  datePersistenceSmokeOnly,
  defaultWorkbook,
  e2eProgressEnabled,
  formulaSmokeOnly,
  formulaUiSmokeOnly,
  gunzipSync,
  macroVbaSubsetOnly,
  mkdtemp,
  nativeShapeWorkbookPath,
  os,
  outlineGroupingOnly,
  pageBackgroundImagePath,
  path,
  pathToFileURL,
  readFile,
  richNativeShapeWorkbookPath,
  saveSmokeOnly,
  smokeOnly,
  standaloneRuntimeCurrentBaseUrl,
  standaloneRuntimeCurrentDir,
  validationMenuOnly,
  viewDiagnosticsOnly,
  viewSmokeOnly,
  workbookArgument,
  workbookPath,
  writeFile,
};
