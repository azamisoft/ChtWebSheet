import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

export const STANDALONE_RUNTIME_VERSION = "0.1.6";
export const STANDALONE_RUNTIME_CHANNEL = "current";

export function standaloneRuntimeRootDir() {
  if (process.platform === "darwin") {
    return path.join(os.homedir(), "Library", "Application Support", "WebSheet", "runtime");
  }

  if (process.platform === "win32") {
    return path.join(process.env.APPDATA || path.join(os.homedir(), "AppData", "Roaming"), "WebSheet", "runtime");
  }

  return path.join(process.env.XDG_DATA_HOME || path.join(os.homedir(), ".local", "share"), "WebSheet", "runtime");
}

export function standaloneRuntimeDir() {
  if (process.env.WEBSHEET_RUNTIME_DIR) {
    return path.resolve(process.env.WEBSHEET_RUNTIME_DIR);
  }
  return path.join(standaloneRuntimeRootDir(), STANDALONE_RUNTIME_VERSION);
}

export function standaloneRuntimeCurrentDir() {
  if (process.env.WEBSHEET_RUNTIME_CURRENT_DIR) {
    return path.resolve(process.env.WEBSHEET_RUNTIME_CURRENT_DIR);
  }
  return path.join(standaloneRuntimeRootDir(), STANDALONE_RUNTIME_CHANNEL);
}

export function standaloneRuntimeBaseUrl() {
  return pathToFileURL(standaloneRuntimeDir()).href;
}

export function standaloneRuntimeCurrentBaseUrl() {
  return pathToFileURL(standaloneRuntimeCurrentDir()).href;
}
