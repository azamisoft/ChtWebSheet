import { chmod, cp, mkdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { STANDALONE_RUNTIME_CHANNEL, STANDALONE_RUNTIME_VERSION } from "./standalone-runtime-path.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const publicRuntimeDir = path.join(projectRoot, "public", "download", "runtime", STANDALONE_RUNTIME_VERSION);
const publicPackagesDir = path.join(projectRoot, "public", "download", "packages");
const distPackagesDir = path.join(projectRoot, "dist", "download", "packages");
const buildRoot = path.join(projectRoot, "build", "runtime-packages");
const macosPackageName = `ChtWebSheet-runtime-${STANDALONE_RUNTIME_VERSION}-macos.zip`;
const windowsPackageName = `ChtWebSheet-runtime-${STANDALONE_RUNTIME_VERSION}-windows.zip`;
const starterHtmlName = "ChtWebSheet.html";
const cwsHtmlFormat = "CWS_HTML";
const cwsSchemaVersion = 1;
const cwsAiGuideUrl = "https://chtec.co.jp/cws/ai/cws-html-guide-v1.json";
const cwsModelSchemaUrl = "https://chtec.co.jp/cws/schema/cws-html-workbook-model-v1.schema.json";
const appIconVersion = "20260612-margin1";
const remoteRuntimeBaseUrl = process.env.WEBSHEET_REMOTE_RUNTIME_BASE_URL ||
  `https://chtec.co.jp/cws/download/runtime/${STANDALONE_RUNTIME_CHANNEL}`;
const versionedRemoteRuntimeBaseUrl = `https://chtec.co.jp/cws/download/runtime/${STANDALONE_RUNTIME_VERSION}`;

await main();

async function main() {
  await mkdir(publicPackagesDir, { recursive: true });
  await rm(buildRoot, { recursive: true, force: true });
  await mkdir(buildRoot, { recursive: true });

  await buildMacosRuntimeZip();
  await buildWindowsRuntimeZip();
  await syncPackagesToDist();

  console.log(`Built runtime packages in ${publicPackagesDir}`);
}

async function buildMacosRuntimeZip() {
  const packageRoot = path.join(buildRoot, "macos", `ChtWebSheet-runtime-${STANDALONE_RUNTIME_VERSION}`);
  const runtimeTarget = path.join(packageRoot, "runtime", STANDALONE_RUNTIME_VERSION);
  await mkdir(runtimeTarget, { recursive: true });
  await cp(publicRuntimeDir, runtimeTarget, { recursive: true, force: true });
  const installPath = path.join(packageRoot, "install.command");
  await writeFile(installPath, macosInstallScript(), "utf8");
  await chmod(installPath, 0o755);
  await writeFile(path.join(packageRoot, starterHtmlName), starterHtml({
    localBaseUrl: `file:///Library/Application%20Support/WebSheet/runtime/${STANDALONE_RUNTIME_CHANNEL}`,
    versionedLocalBaseUrl: `file:///Library/Application%20Support/WebSheet/runtime/${STANDALONE_RUNTIME_VERSION}`,
    title: "Cht WebSheet",
  }), "utf8");
  await writeFile(path.join(packageRoot, "README.txt"), macosReadme(), "utf8");
  await rm(path.join(publicPackagesDir, `ChtWebSheet-${STANDALONE_RUNTIME_VERSION}.pkg`), { force: true });
  await rm(path.join(publicPackagesDir, macosPackageName), { force: true });
  await run("zip", ["-qr", path.join(publicPackagesDir, macosPackageName), "."], { cwd: path.dirname(packageRoot) });
}

async function buildWindowsRuntimeZip() {
  const packageRoot = path.join(buildRoot, "windows", `ChtWebSheet-runtime-${STANDALONE_RUNTIME_VERSION}`);
  const runtimeTarget = path.join(packageRoot, "runtime", STANDALONE_RUNTIME_VERSION);
  await mkdir(runtimeTarget, { recursive: true });
  await cp(publicRuntimeDir, runtimeTarget, { recursive: true, force: true });
  await writeFile(path.join(packageRoot, "install-runtime.ps1"), windowsInstallScript(), "utf8");
  await writeFile(path.join(packageRoot, starterHtmlName), starterHtml({
    localBaseUrl: `file:///C:/ProgramData/WebSheet/runtime/${STANDALONE_RUNTIME_CHANNEL}`,
    versionedLocalBaseUrl: `file:///C:/ProgramData/WebSheet/runtime/${STANDALONE_RUNTIME_VERSION}`,
    title: "Cht WebSheet",
  }), "utf8");
  await writeFile(path.join(packageRoot, "README.txt"), windowsReadme(), "utf8");
  await rm(path.join(publicPackagesDir, windowsPackageName), { force: true });
  await run("zip", ["-qr", path.join(publicPackagesDir, windowsPackageName), "."], { cwd: path.dirname(packageRoot) });
}

async function syncPackagesToDist() {
  await mkdir(distPackagesDir, { recursive: true });
  await cp(publicPackagesDir, distPackagesDir, { recursive: true, force: true });
}

function windowsInstallScript() {
  return `# Cht WebSheet common runtime installer
$ErrorActionPreference = "Stop"
$version = "${STANDALONE_RUNTIME_VERSION}"
$source = Join-Path $PSScriptRoot "runtime\\$version"
$runtimeRoot = Join-Path $env:ProgramData "WebSheet\\runtime"
$versionTarget = Join-Path $runtimeRoot $version
$currentTarget = Join-Path $runtimeRoot "${STANDALONE_RUNTIME_CHANNEL}"
New-Item -ItemType Directory -Force -Path $versionTarget | Out-Null
Copy-Item -Path (Join-Path $source "*") -Destination $versionTarget -Recurse -Force
if (Test-Path $currentTarget) { Remove-Item $currentTarget -Recurse -Force }
New-Item -ItemType Directory -Force -Path $currentTarget | Out-Null
Copy-Item -Path (Join-Path $source "*") -Destination $currentTarget -Recurse -Force
Write-Host "Cht WebSheet common runtime $version installed to $versionTarget"
Write-Host "Stable runtime entry updated: $currentTarget"
`;
}

function macosInstallScript() {
  return `#!/bin/bash
set -euo pipefail

VERSION="${STANDALONE_RUNTIME_VERSION}"
SOURCE_DIR="$(cd "$(dirname "$0")" && pwd)/runtime/$VERSION"
RUNTIME_ROOT="/Library/Application Support/WebSheet/runtime"
VERSION_TARGET_DIR="$RUNTIME_ROOT/$VERSION"
CURRENT_TARGET_DIR="$RUNTIME_ROOT/${STANDALONE_RUNTIME_CHANNEL}"

if [ ! -d "$SOURCE_DIR" ]; then
  echo "Runtime source not found: $SOURCE_DIR" >&2
  exit 1
fi

echo "Cht WebSheet common runtime $VERSION"
echo "Installing to: $VERSION_TARGET_DIR"
echo "Updating stable entry: $CURRENT_TARGET_DIR"
echo "管理者パスワードを求められた場合は、macOS の確認に従って入力してください。"

sudo mkdir -p "$VERSION_TARGET_DIR"
sudo cp -R "$SOURCE_DIR"/. "$VERSION_TARGET_DIR"/
sudo rm -rf "$CURRENT_TARGET_DIR"
sudo mkdir -p "$CURRENT_TARGET_DIR"
sudo cp -R "$SOURCE_DIR"/. "$CURRENT_TARGET_DIR"/
sudo chmod -R a+rX "/Library/Application Support/WebSheet"

echo ""
echo "Installation complete."
echo "You can close this window."
`;
}

function macosReadme() {
  return `Cht WebSheet common runtime ${STANDALONE_RUNTIME_VERSION}

Install on macOS:
1. Extract this zip file.
2. Double-click install.command.
3. If macOS asks for permission or an administrator password, follow the prompt.
4. The runtime will be installed to:
   /Library/Application Support/WebSheet/runtime/${STANDALONE_RUNTIME_CHANNEL}
   /Library/Application Support/WebSheet/runtime/<version>
   where <version> is the runtime version included in this package.

After installation, Cht WebSheet HTML workbooks can load the local runtime offline.
Saved HTML workbooks use the stable "runtime/${STANDALONE_RUNTIME_CHANNEL}" entry first, so runtime updates do not require editing old files.
You can also open ChtWebSheet.html in this package to start a blank CWS workbook.
`;
}

function windowsReadme() {
  return `Cht WebSheet common runtime ${STANDALONE_RUNTIME_VERSION}

Install on Windows:
1. Extract this zip file.
2. Right-click install-runtime.ps1 and run it with PowerShell.
3. The runtime will be installed to:
   C:\\ProgramData\\WebSheet\\runtime\\${STANDALONE_RUNTIME_CHANNEL}
   C:\\ProgramData\\WebSheet\\runtime\\<version>
   where <version> is the runtime version included in this package.

After installation, Cht WebSheet HTML workbooks can load the local runtime offline.
Saved HTML workbooks use the stable "runtime\\${STANDALONE_RUNTIME_CHANNEL}" entry first, so runtime updates do not require editing old files.
You can also open ChtWebSheet.html in this package to start a blank CWS workbook.
`;
}

function starterHtml({ localBaseUrl, versionedLocalBaseUrl, title }) {
  return `<!doctype html>
<html lang="ja" data-cws-format="${cwsHtmlFormat}" data-cws-schema-version="${cwsSchemaVersion}" data-cws-app="CWS">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<meta name="generator" content="Cht WebSheet ${STANDALONE_RUNTIME_VERSION}" />
<meta name="cws:format" content="${cwsHtmlFormat}" />
<meta name="cws:schema-version" content="${cwsSchemaVersion}" />
<meta name="cws:guide" content="${cwsAiGuideUrl}" />
<meta name="cws:schema" content="${cwsModelSchemaUrl}" />
<link rel="icon" type="image/png" href="runtime/${STANDALONE_RUNTIME_VERSION}/icon.png?v=${appIconVersion}" />
<title>${escapeHtml(title)}</title>
</head>
<body>
<div id="websheetStandaloneBoot" style="font:14px system-ui,sans-serif;padding:16px;color:#1f2937">
  <strong>${escapeHtml(title)}</strong>
  <div>Cht WebSheet 共通 runtime を読み込み中...</div>
</div>
<script>
${starterLoaderScript({
  version: STANDALONE_RUNTIME_VERSION,
  localBaseUrl,
  versionedLocalBaseUrl,
  remoteBaseUrl: remoteRuntimeBaseUrl,
  versionedRemoteBaseUrl: versionedRemoteRuntimeBaseUrl,
})}
</script>
</body>
</html>
`;
}

function starterLoaderScript(config) {
  return `(function(){
  "use strict";
  var config = ${JSON.stringify(config).replace(/</g, "\\u003c")};
  var remoteRuntimeCacheBuster = Date.now().toString(36);
  function baseUrl(value){ return String(value || "").replace(/\\/+$/, ""); }
  function asset(base, file){ return baseUrl(base) + "/" + file; }
  function runtimeAsset(base, file, source){
    var url = asset(base, file);
    if (source !== "remote") return url;
    return url + (url.indexOf("?") >= 0 ? "&" : "?") + "v=" + encodeURIComponent(config.version || "") + "&b=" + remoteRuntimeCacheBuster;
  }
  function status(message){
    var boot = document.getElementById("websheetStandaloneBoot");
    if (!boot) return;
    var detail = boot.querySelector("div") || boot;
    detail.textContent = message;
  }
  function loadCss(href){
    return new Promise(function(resolve, reject){
      var link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = href;
      link.onload = resolve;
      link.onerror = function(){ reject(new Error("スタイルシートを読み込めませんでした: " + href)); };
      document.head.appendChild(link);
    });
  }
  function loadScript(src){
    return new Promise(function(resolve, reject){
      var script = document.createElement("script");
      script.src = src;
      script.async = false;
      script.onload = resolve;
      script.onerror = function(){ reject(new Error("スクリプトを読み込めませんでした: " + src)); };
      document.head.appendChild(script);
    });
  }
  function probeRuntimeVersion(base){
    return new Promise(function(resolve){
      if (!base) {
        resolve(null);
        return;
      }
      var done = false;
      var previousProbe = window.__WEBSHEET_RUNTIME_VERSION_PROBE__;
      var timer = setTimeout(function(){ finish(null); }, 1200);
      function finish(info){
        if (done) return;
        done = true;
        clearTimeout(timer);
        window.__WEBSHEET_RUNTIME_VERSION_PROBE__ = previousProbe;
        resolve(info || window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__ || null);
      }
      window.__WEBSHEET_RUNTIME_VERSION_PROBE__ = finish;
      var script = document.createElement("script");
      script.src = asset(base, "version.js");
      script.async = true;
      script.onload = function(){ finish(window.__WEBSHEET_STANDALONE_RUNTIME_VERSION__ || null); };
      script.onerror = function(){ finish(null); };
      document.head.appendChild(script);
    });
  }
  function unique(values){
    var result = [];
    values.forEach(function(value){
      value = baseUrl(value);
      if (value && result.indexOf(value) < 0) result.push(value);
    });
    return result;
  }
  function versionParts(value){
    return String(value || "").split(/[.-]/).map(function(part){
      var number = parseInt(part, 10);
      return Number.isFinite(number) ? number : 0;
    });
  }
  function compareVersions(left, right){
    var a = versionParts(left);
    var b = versionParts(right);
    var length = Math.max(a.length, b.length, 3);
    for (var index = 0; index < length; index += 1) {
      var delta = (a[index] || 0) - (b[index] || 0);
      if (delta) return delta;
    }
    return 0;
  }
  function runtimeIsUsable(info){
    return !!(info && info.version && compareVersions(info.version, config.version) >= 0);
  }
  async function firstUsableLocalRuntime(){
    var firstInfo = null;
    var firstBase = "";
    var candidates = unique([config.localBaseUrl, config.versionedLocalBaseUrl]);
    for (var index = 0; index < candidates.length; index += 1) {
      var base = candidates[index];
      var info = await probeRuntimeVersion(base);
      if (info && !firstInfo) {
        firstInfo = info;
        firstBase = base;
      }
      if (runtimeIsUsable(info)) {
        return { base: base, info: info, status: "current" };
      }
    }
    return { base: firstBase, info: firstInfo, status: firstInfo ? "outdated" : "missing" };
  }
  async function loadRuntime(base, source){
    window.__WEBSHEET_RUNTIME_SOURCE__ = source;
    status(source === "local" ? "ローカル共通 runtime を読み込み中..." : "共通 runtime をダウンロード中...");
    await loadCss(runtimeAsset(base, "websheet-standalone.css", source));
    await loadScript(runtimeAsset(base, "websheet-standalone.js", source));
  }
  async function boot(){
    var localRuntime = await firstUsableLocalRuntime();
    window.__WEBSHEET_LOCAL_RUNTIME_INFO__ = localRuntime.info || null;
    window.__WEBSHEET_LOCAL_RUNTIME_STATUS__ = localRuntime.status;
    if (localRuntime.status === "current") {
      await loadRuntime(localRuntime.base, "local");
      return;
    }
    var remoteCandidates = unique([config.remoteBaseUrl, config.versionedRemoteBaseUrl]);
    var lastError = null;
    for (var index = 0; index < remoteCandidates.length; index += 1) {
      try {
        await loadRuntime(remoteCandidates[index], "remote");
        return;
      } catch (error) {
        lastError = error;
      }
    }
    throw lastError || new Error("共通 runtime を読み込めませんでした。");
  }
  boot().catch(function(error){
    console.error(error);
    status("Cht WebSheet 共通 runtime を読み込めませんでした。");
    var errorBox = document.createElement("pre");
    errorBox.style.cssText = "margin:16px;padding:12px;border:1px solid #ef4444;background:#fff1f2;color:#991b1b;white-space:pre-wrap";
    errorBox.textContent = "Cht WebSheet 共通 runtime を読み込めませんでした。\\n" + (error && error.message ? error.message : String(error));
    document.body.appendChild(errorBox);
  });
})();`;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: options.cwd || projectRoot,
      stdio: "inherit",
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${command} exited with code ${code}`));
    });
  });
}
