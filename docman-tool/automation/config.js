const fs = require("fs");
const path = require("path");

function findProjectRoot(startDir) {
  let dir = startDir;
  while (true) {
    if (fs.existsSync(path.join(dir, "run.js"))) return dir;
    const candidate = path.join(dir, "package.json");
    if (fs.existsSync(candidate)) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) return startDir;
    dir = parent;
  }
}

const PROJECT_ROOT = findProjectRoot(path.resolve(__dirname, ".."));
const DEFAULT_CONFIG_PATH = path.join(PROJECT_ROOT, "docman-tool.config.json");

const DEFAULT_CONFIG = {
  modeDefaults: {
    defaultMode: "login",
  },
  browser: {
    window: {
      width: 1440,
      height: 900,
      positionX: 40,
      positionY: 40,
    },
    preDocmanHeadless: true,
    allowVisibleFallback: false,
    step4Visible: true,
    step4UseCurrentChrome: true,
    step4BrowserEngine: "chrome",
    chromeCdpUrl: "http://127.0.0.1:9222",
  },
  betterLetter: {
    basicAuth: {
      username: "",
      password: "",
    },
  },
  clean: {
    batchSize: 50,
    defaultSourceFolder: "",
    defaultDestinationFolder: "",
    folderPicker: {
      enabled: true,
      maxSuggestions: 10,
      maxScrollPasses: 90,
    },
  },
  retries: {
    step: {
      attempts: 2,
      baseDelayMs: 500,
      maxDelayMs: 3000,
    },
  },
  timeouts: {
    navigationMs: 60000,
    selectorMs: 60000,
    docmanLoginFastCheckMs: 3500,
    docmanLoginDeepCheckMs: 9000,
  },
  logging: {
    enabled: true,
    directory: "logs",
  },
};

function deepMerge(base, extra) {
  if (!extra || typeof extra !== "object") return clone(base);
  const out = Array.isArray(base) ? [...base] : { ...base };

  for (const [key, value] of Object.entries(extra)) {
    if (
      value &&
      typeof value === "object" &&
      !Array.isArray(value) &&
      out[key] &&
      typeof out[key] === "object" &&
      !Array.isArray(out[key])
    ) {
      out[key] = deepMerge(out[key], value);
      continue;
    }
    out[key] = clone(value);
  }

  return out;
}

function clone(value) {
  if (Array.isArray(value)) return value.map((v) => clone(v));
  if (value && typeof value === "object") {
    const out = {};
    for (const [k, v] of Object.entries(value)) out[k] = clone(v);
    return out;
  }
  return value;
}

function ensureInteger(config, pathLabel, min, max) {
  const value = pathLabel
    .split(".")
    .reduce((acc, key) => (acc == null ? acc : acc[key]), config);
  if (!Number.isInteger(value) || value < min || value > max) {
    throw new Error(
      `Invalid config at "${pathLabel}": expected integer between ${min} and ${max}, got ${value}`
    );
  }
}

function ensureBoolean(config, pathLabel) {
  const value = pathLabel
    .split(".")
    .reduce((acc, key) => (acc == null ? acc : acc[key]), config);
  if (typeof value !== "boolean") {
    throw new Error(`Invalid config at "${pathLabel}": expected boolean, got ${typeof value}`);
  }
}

function ensureString(config, pathLabel) {
  const value = pathLabel
    .split(".")
    .reduce((acc, key) => (acc == null ? acc : acc[key]), config);
  if (typeof value !== "string") {
    throw new Error(`Invalid config at "${pathLabel}": expected string, got ${typeof value}`);
  }
}

function validateConfig(config) {
  const mode = config?.modeDefaults?.defaultMode;
  if (!["login", "verify", "create-group", "clean", "onboarding"].includes(mode)) {
    throw new Error(
      `Invalid config at "modeDefaults.defaultMode": expected one of login|verify|create-group|clean|onboarding, got ${mode}`
    );
  }

  ensureInteger(config, "browser.window.width", 900, 6000);
  ensureInteger(config, "browser.window.height", 600, 4000);
  ensureInteger(config, "browser.window.positionX", 0, 5000);
  ensureInteger(config, "browser.window.positionY", 0, 5000);
  ensureBoolean(config, "browser.preDocmanHeadless");
  ensureBoolean(config, "browser.allowVisibleFallback");
  ensureBoolean(config, "browser.step4Visible");
  ensureBoolean(config, "browser.step4UseCurrentChrome");
  ensureString(config, "browser.step4BrowserEngine");
  ensureString(config, "browser.chromeCdpUrl");
  if (!["chrome", "chromium"].includes(String(config.browser.step4BrowserEngine).toLowerCase())) {
    throw new Error(
      `Invalid config at "browser.step4BrowserEngine": expected chrome|chromium, got ${config.browser.step4BrowserEngine}`
    );
  }

  ensureInteger(config, "clean.batchSize", 1, 500);
  ensureString(config, "clean.defaultSourceFolder");
  ensureString(config, "clean.defaultDestinationFolder");
  ensureBoolean(config, "clean.folderPicker.enabled");
  ensureInteger(config, "clean.folderPicker.maxSuggestions", 3, 25);
  ensureInteger(config, "clean.folderPicker.maxScrollPasses", 10, 400);

  ensureInteger(config, "retries.step.attempts", 1, 10);
  ensureInteger(config, "retries.step.baseDelayMs", 0, 60000);
  ensureInteger(config, "retries.step.maxDelayMs", 0, 120000);
  if (config.retries.step.maxDelayMs < config.retries.step.baseDelayMs) {
    throw new Error(
      "Invalid config at \"retries.step.maxDelayMs\": must be >= retries.step.baseDelayMs"
    );
  }

  ensureInteger(config, "timeouts.navigationMs", 5000, 180000);
  ensureInteger(config, "timeouts.selectorMs", 5000, 180000);
  ensureInteger(config, "timeouts.docmanLoginFastCheckMs", 500, 30000);
  ensureInteger(config, "timeouts.docmanLoginDeepCheckMs", 1000, 120000);

  ensureBoolean(config, "logging.enabled");
  ensureString(config, "logging.directory");
  ensureString(config, "betterLetter.basicAuth.username");
  ensureString(config, "betterLetter.basicAuth.password");
}

function loadRuntimeConfig(configPath = DEFAULT_CONFIG_PATH) {
  const exists = fs.existsSync(configPath);
  let fileConfig = {};

  if (exists) {
    const raw = fs.readFileSync(configPath, "utf8");
    try {
      fileConfig = JSON.parse(raw);
    } catch (error) {
      throw new Error(
        `Failed to parse config file at ${configPath}: ${error?.message || "invalid JSON"}`
      );
    }
  }

  const config = deepMerge(DEFAULT_CONFIG, fileConfig);
  validateConfig(config);

  return {
    config,
    configPath,
    hasConfigFile: exists,
    projectRoot: PROJECT_ROOT,
  };
}

module.exports = {
  DEFAULT_CONFIG,
  DEFAULT_CONFIG_PATH,
  loadRuntimeConfig,
};
