const CWS_AGENT_HOST_VERSION = 1;
const CWS_AGENT_DEFAULT_REMOTE_BASE_URL = "https://chtec.co.jp/cws/cws-agent/";
const CWS_AGENT_ASSET_VERSION = "20260619-streaming-agent-4";

let cwsAgentLoadStarted = false;

export function initCwsAgentBridge(runtimeOptions = {}) {
  if (typeof window === "undefined" || typeof document === "undefined") return;

  const agentBaseUrl = cwsAgentBaseUrl();
  const apiBaseUrl = cwsAgentApiBaseUrl(agentBaseUrl);
  const runtime = {
    ...runtimeOptions,
    agentBaseUrl,
    apiBaseUrl,
  };

  const host = window.ChtWebSheetAgentHost || {};
  host.version = CWS_AGENT_HOST_VERSION;
  host.runtime = runtime;
  host.register = (agent) => {
    if (!agent || typeof agent.init !== "function") return;
    host.agent = agent;
    if (host.initialized && host.initializedRuntime === host.runtime) return;
    host.initialized = true;
    host.initializedRuntime = host.runtime;
    try {
      agent.init(host.runtime);
    } catch (error) {
      console.warn("CWS Agent initialization failed:", error);
      host.initialized = false;
      host.initializedRuntime = null;
    }
  };
  window.ChtWebSheetAgentHost = host;
  window.ChtWebSheetAgentRuntime = runtime;

  if (window.ChtWebSheetAgent) {
    host.register(window.ChtWebSheetAgent);
    return;
  }
  loadCwsAgentAssets(agentBaseUrl);
}

function loadCwsAgentAssets(agentBaseUrl) {
  if (!agentBaseUrl || cwsAgentLoadStarted) return;
  cwsAgentLoadStarted = true;
  loadCwsAgentStylesheet(cwsAgentAssetUrl("agent.css", agentBaseUrl));
  const script = document.createElement("script");
  script.src = cwsAgentAssetUrl("agent.js", agentBaseUrl);
  script.async = true;
  script.dataset.cwsAgentScript = "true";
  script.onerror = () => {
    cwsAgentLoadStarted = false;
    console.info("CWS Agent is not available at this location.");
  };
  document.head.appendChild(script);
}

function cwsAgentAssetUrl(assetName, agentBaseUrl) {
  const url = new URL(assetName, agentBaseUrl);
  url.searchParams.set("v", CWS_AGENT_ASSET_VERSION);
  return url.href;
}

function loadCwsAgentStylesheet(href) {
  if (!href) return;
  const exists = Array.from(document.querySelectorAll("link[data-cws-agent-css]"))
    .some((link) => link.href === href);
  if (exists) return;
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = href;
  link.dataset.cwsAgentCss = "true";
  document.head.appendChild(link);
}

function cwsAgentBaseUrl() {
  const explicit = cwsAgentExplicitBaseUrl();
  if (explicit) return explicit;
  if (location.protocol === "file:" || location.origin === "null") {
    return CWS_AGENT_DEFAULT_REMOTE_BASE_URL;
  }
  const baseUrl = cwsAgentLocationBaseUrl();
  if (!baseUrl) return CWS_AGENT_DEFAULT_REMOTE_BASE_URL;
  const pathname = baseUrl.pathname || "/";
  const cwsMatch = pathname.match(/^(.*\/cws)(?:\/|$)/);
  if (cwsMatch) return `${baseUrl.origin}${cwsMatch[1]}/cws-agent/`;
  if (location.protocol === "blob:") return CWS_AGENT_DEFAULT_REMOTE_BASE_URL;
  return new URL("./cws-agent/", baseUrl.href).href;
}

function cwsAgentApiBaseUrl(agentBaseUrl) {
  const explicit = cwsAgentExplicitApiBaseUrl();
  if (explicit) return explicit;
  return new URL("api/", agentBaseUrl || CWS_AGENT_DEFAULT_REMOTE_BASE_URL).href;
}

function cwsAgentExplicitBaseUrl() {
  return normalizeCwsAgentUrl(
    window.__CWS_AGENT_BASE_URL__ ||
      document.querySelector('meta[name="cws-agent-base"]')?.content ||
      "",
  );
}

function cwsAgentExplicitApiBaseUrl() {
  return normalizeCwsAgentUrl(
    window.__CWS_AGENT_API_BASE_URL__ ||
      document.querySelector('meta[name="cws-agent-api-base"]')?.content ||
      "",
  );
}

function normalizeCwsAgentUrl(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  try {
    const url = new URL(text, cwsAgentLocationBaseUrl()?.href || CWS_AGENT_DEFAULT_REMOTE_BASE_URL);
    return url.href.endsWith("/") ? url.href : `${url.href}/`;
  } catch {
    return "";
  }
}

function cwsAgentLocationBaseUrl() {
  try {
    if (location.protocol === "blob:") {
      const innerUrl = new URL(String(location.href || "").replace(/^blob:/, ""));
      if (/^https?:$/i.test(innerUrl.protocol)) return innerUrl;
      return null;
    }
    const url = new URL(location.href);
    return /^https?:$/i.test(url.protocol) ? url : null;
  } catch {
    return null;
  }
}
