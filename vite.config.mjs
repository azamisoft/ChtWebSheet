import { access, readFile, realpath, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { defineConfig } from "vite";
import { standaloneRuntimeCurrentBaseUrl } from "./scripts/standalone-runtime-path.mjs";

const LOCAL_SAVE_ROUTE = "/__websheet/local-save";
const LOCAL_OPEN_WORKBOOK_ROUTE = "/__websheet/local-open-workbook";
const MAX_LOCAL_SAVE_BYTES = 120 * 1024 * 1024;

function webSheetLocalSavePlugin() {
  return {
    name: "websheet-local-save",
    configureServer(server) {
      server.middlewares.use(LOCAL_SAVE_ROUTE, async (req, res, next) => {
        if (req.method !== "POST" && req.method !== "OPTIONS") {
          next();
          return;
        }

        setLocalSaveCorsHeaders(req, res);
        if (req.method === "OPTIONS") {
          res.statusCode = 204;
          res.end();
          return;
        }

        try {
          const body = await readJsonBody(req);
          const result = await saveStandaloneHtmlToExistingFile(body);
          sendJson(res, 200, result);
        } catch (error) {
          const status = Number(error?.status) || 500;
          sendJson(res, status, { ok: false, message: error?.message || String(error) });
        }
      });
      server.middlewares.use(LOCAL_OPEN_WORKBOOK_ROUTE, async (req, res, next) => {
        if (req.method !== "POST" && req.method !== "OPTIONS") {
          next();
          return;
        }

        setLocalSaveCorsHeaders(req, res);
        if (req.method === "OPTIONS") {
          res.statusCode = 204;
          res.end();
          return;
        }

        try {
          const body = await readJsonBody(req);
          const result = await createStandaloneHtmlForDroppedWorkbook(body);
          sendJson(res, 200, result);
        } catch (error) {
          const status = Number(error?.status) || 500;
          sendJson(res, status, { ok: false, message: error?.message || String(error) });
        }
      });
    },
  };
}

function setLocalSaveCorsHeaders(req, res) {
  const origin = String(req.headers.origin || "");
  if (!origin || origin === "null" || origin === "file://" || /^https?:\/\/(?:127\.0\.0\.1|localhost)(?::\d+)?$/i.test(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin || "*");
  }
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (String(req.headers["access-control-request-private-network"] || "").toLowerCase() === "true") {
    res.setHeader("Access-Control-Allow-Private-Network", "true");
  }
  res.setHeader("Access-Control-Max-Age", "600");
}

async function readJsonBody(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > MAX_LOCAL_SAVE_BYTES) throw httpError(413, "保存データが大きすぎます。");
    chunks.push(chunk);
  }
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}");
  } catch {
    throw httpError(400, "保存リクエストを解析できません。");
  }
}

async function saveStandaloneHtmlToExistingFile(body) {
  const rawPath = String(body?.path || "");
  const html = String(body?.html || "");
  if (!path.isAbsolute(rawPath)) throw httpError(400, "保存先ファイルを特定できません。");
  if (!isWebSheetStandaloneHtml(html)) throw httpError(400, "WebSheet HTML 以外は保存できません。");

  const targetPath = await realpath(rawPath).catch((error) => {
    if (isPermissionError(error)) {
      throw httpError(403, "保存先ファイルへのアクセス権がありません。Downloads など保護されたフォルダーに保存する場合は、WebSheet dev server を実行しているアプリにアクセス権を付与してください。");
    }
    return "";
  });
  if (!targetPath) throw httpError(404, "保存先ファイルが見つかりません。");
  const ext = path.extname(targetPath).toLowerCase();
  if (ext !== ".html" && ext !== ".htm") throw httpError(400, "HTML ファイルだけを上書きできます。");

  const previous = await readFile(targetPath, "utf8").catch((error) => {
    if (isPermissionError(error)) {
      throw httpError(403, "保存先ファイルを読み取る権限がありません。WebSheet dev server を実行しているアプリにアクセス権を付与してください。");
    }
    return "";
  });
  if (!isWebSheetStandaloneHtml(previous)) {
    throw httpError(403, "既存の WebSheet HTML ファイルだけを上書きできます。");
  }

  await writeFile(targetPath, html, "utf8").catch((error) => {
    if (isPermissionError(error)) {
      throw httpError(403, "保存先ファイルに書き込む権限がありません。WebSheet dev server を実行しているアプリにアクセス権を付与してください。");
    }
    throw error;
  });
  return { ok: true, path: targetPath };
}

async function createStandaloneHtmlForDroppedWorkbook(body) {
  const rawCurrentPath = String(body?.currentPath || "");
  const fileName = sanitizeOutputHtmlFileName(body?.fileName || "workbook.html");
  const html = String(body?.html || "");
  if (!path.isAbsolute(rawCurrentPath)) throw httpError(400, "現在の HTML ファイルを特定できません。");
  if (!isWebSheetStandaloneHtml(html)) throw httpError(400, "WebSheet HTML 以外は作成できません。");

  const currentPath = await realpath(rawCurrentPath).catch((error) => {
    if (isPermissionError(error)) {
      throw httpError(403, "現在の HTML ファイルへのアクセス権がありません。WebSheet dev server を実行しているアプリにアクセス権を付与してください。");
    }
    return "";
  });
  if (!currentPath) throw httpError(404, "現在の HTML ファイルが見つかりません。");
  const currentHtml = await readFile(currentPath, "utf8").catch((error) => {
    if (isPermissionError(error)) {
      throw httpError(403, "現在の HTML ファイルを読み取る権限がありません。WebSheet dev server を実行しているアプリにアクセス権を付与してください。");
    }
    return "";
  });
  if (!isWebSheetStandaloneHtml(currentHtml)) {
    throw httpError(403, "WebSheet HTML を開いている場合だけ、新しいローカル HTML を作成できます。");
  }

  const targetPath = await uniqueHtmlPath(path.dirname(currentPath), fileName);
  await writeFile(targetPath, html, { encoding: "utf8", flag: "wx" }).catch((error) => {
    if (isPermissionError(error)) {
      throw httpError(403, "新しい HTML ファイルに書き込む権限がありません。WebSheet dev server を実行しているアプリにアクセス権を付与してください。");
    }
    throw error;
  });
  return { ok: true, path: targetPath, url: pathToFileURL(targetPath).href };
}

function isWebSheetStandaloneHtml(value) {
  const text = String(value || "");
  return text.includes('id="websheet-model"') && /websheet/i.test(text);
}

function sanitizeOutputHtmlFileName(value) {
  const raw = String(value || "").trim().split(/[\\/]/).pop() || "workbook.html";
  const cleaned = raw
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_")
    .replace(/\s+$/g, "")
    .replace(/^\.+$/, "workbook");
  const withExtension = path.extname(cleaned).toLowerCase() === ".html" || path.extname(cleaned).toLowerCase() === ".htm"
    ? cleaned
    : `${cleaned.replace(/\.[^.]*$/, "") || "workbook"}.html`;
  return withExtension || "workbook.html";
}

async function uniqueHtmlPath(dir, fileName) {
  const parsed = path.parse(fileName);
  const base = parsed.name || "workbook";
  const ext = parsed.ext || ".html";
  for (let index = 0; index < 1000; index += 1) {
    const candidate = path.join(dir, index === 0 ? `${base}${ext}` : `${base} (${index})${ext}`);
    const exists = await access(candidate).then(() => true, () => false);
    if (!exists) return candidate;
  }
  throw httpError(409, "作成できるファイル名を決定できません。");
}

function sendJson(res, status, payload) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(payload));
}

function httpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

function isPermissionError(error) {
  return error?.code === "EPERM" || error?.code === "EACCES";
}

export default defineConfig({
  define: {
    __WEBSHEET_DEFAULT_RUNTIME_BASE_URL__: JSON.stringify(
      process.env.WEBSHEET_DEFAULT_RUNTIME_BASE_URL || standaloneRuntimeCurrentBaseUrl(),
    ),
  },
  server: {
    cors: false,
  },
  plugins: [webSheetLocalSavePlugin()],
});
