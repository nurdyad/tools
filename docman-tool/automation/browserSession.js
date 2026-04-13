// automation/browserSession.js
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");

function findProjectRoot(startDir) {
  let dir = startDir;
  while (true) {
    const runCandidate = path.join(dir, "run.js");
    if (fs.existsSync(runCandidate)) return dir;
    const candidate = path.join(dir, "package.json");
    if (fs.existsSync(candidate)) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) return startDir;
    dir = parent;
  }
}

const PROJECT_ROOT = findProjectRoot(path.resolve(__dirname, ".."));
const PROFILE_DIR = path.join(PROJECT_ROOT, ".browser-profile");
const LOCK_FILE = path.join(PROFILE_DIR, "playwright-profile.lock");
const DEFAULT_WINDOW = {
  width: 1440,
  height: 900,
  positionX: 40,
  positionY: 40,
};
const DEFAULT_CHROME_CDP_URL = "http://127.0.0.1:9222";

let _session = null;
let _sessionPromise = null;
let _lockAcquired = false;
let _signalsRegistered = false;
let _closingByScript = false;
let _manualCloseExitPending = false;
let _lastLaunchOptions = null;

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function acquireLock() {
  ensureDir(PROFILE_DIR);

  if (fs.existsSync(LOCK_FILE)) {
    const lockContent = fs.readFileSync(LOCK_FILE, "utf8").trim();
    const lockInfo = parseLockInfo(lockContent);

    if (lockInfo?.pid && !isProcessAlive(lockInfo.pid)) {
      console.warn(
        `[browserSession] Removing stale profile lock from pid ${lockInfo.pid}: ${LOCK_FILE}`
      );
      cleanupStaleProfileArtifacts();
    }
  }

  if (fs.existsSync(LOCK_FILE)) {
    const lockContent = fs.readFileSync(LOCK_FILE, "utf8").trim();
    throw new Error(
      [
        `Profile appears to be in use (lock file exists): ${LOCK_FILE}`,
        lockContent ? `Lock info: ${lockContent}` : "",
        "Close any other running instance of this automation (or Chromium using this profile) and try again.",
      ]
        .filter(Boolean)
        .join("\n")
    );
  }

  fs.writeFileSync(
    LOCK_FILE,
    `pid=${process.pid}\nstarted=${new Date().toISOString()}\nprofile=${PROFILE_DIR}\n`,
    "utf8"
  );
  _lockAcquired = true;
}

function releaseLock() {
  try {
    if (fs.existsSync(LOCK_FILE)) fs.unlinkSync(LOCK_FILE);
  } catch (_) {}
  _lockAcquired = false;
}

function parseLockInfo(lockContent = "") {
  const lines = String(lockContent || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const info = {};

  for (const line of lines) {
    const eqIndex = line.indexOf("=");
    if (eqIndex === -1) continue;
    const key = line.slice(0, eqIndex).trim();
    const value = line.slice(eqIndex + 1).trim();
    if (!key) continue;
    info[key] = value;
  }

  const pid = Number.parseInt(info.pid, 10);
  return {
    ...info,
    pid: Number.isInteger(pid) && pid > 0 ? pid : null,
  };
}

function isProcessAlive(pid) {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    if (error?.code === "ESRCH") return false;
    if (error?.code === "EPERM") return true;
    return true;
  }
}

function cleanupStaleProfileArtifacts() {
  const staleTargets = [
    LOCK_FILE,
    path.join(PROFILE_DIR, "SingletonLock"),
    path.join(PROFILE_DIR, "SingletonSocket"),
    path.join(PROFILE_DIR, "SingletonCookie"),
    path.join(PROFILE_DIR, "RunningChromeVersion"),
  ];

  for (const target of staleTargets) {
    try {
      if (fs.existsSync(target)) {
        fs.rmSync(target, { force: true });
      }
    } catch (_) {}
  }
}

function patchJsonFile(filePath, patcher) {
  try {
    if (!fs.existsSync(filePath)) return false;
    const raw = fs.readFileSync(filePath, "utf8");
    if (!raw.trim()) return false;
    const parsed = JSON.parse(raw);
    const changed = patcher(parsed);
    if (!changed) return false;
    fs.writeFileSync(filePath, JSON.stringify(parsed), "utf8");
    return true;
  } catch (_) {
    return false;
  }
}

function markChromiumProfileAsClean() {
  const prefPath = path.join(PROFILE_DIR, "Default", "Preferences");
  patchJsonFile(prefPath, (json) => {
    let changed = false;
    if (typeof json.exited_cleanly === "boolean" && json.exited_cleanly !== true) {
      json.exited_cleanly = true;
      changed = true;
    }
    if (typeof json.exit_type === "string" && json.exit_type.toLowerCase() !== "normal") {
      json.exit_type = "Normal";
      changed = true;
    }
    if (json.profile && typeof json.profile === "object") {
      if (json.profile.exited_cleanly !== true) {
        json.profile.exited_cleanly = true;
        changed = true;
      }
      if (json.profile.exit_type !== "Normal") {
        json.profile.exit_type = "Normal";
        changed = true;
      }
    }
    return changed;
  });

  const localStatePath = path.join(PROFILE_DIR, "Local State");
  patchJsonFile(localStatePath, (json) => {
    const stability = json?.user_experience_metrics?.stability;
    if (!stability || typeof stability !== "object") return false;
    if (stability.exited_cleanly === true) return false;
    stability.exited_cleanly = true;
    return true;
  });
}

function resolveWindowOptions(windowOptions = {}) {
  const width = Number.isInteger(windowOptions.width) ? windowOptions.width : DEFAULT_WINDOW.width;
  const height = Number.isInteger(windowOptions.height)
    ? windowOptions.height
    : DEFAULT_WINDOW.height;
  const positionX = Number.isInteger(windowOptions.positionX)
    ? windowOptions.positionX
    : DEFAULT_WINDOW.positionX;
  const positionY = Number.isInteger(windowOptions.positionY)
    ? windowOptions.positionY
    : DEFAULT_WINDOW.positionY;

  return {
    width,
    height,
    positionX,
    positionY,
  };
}

function normalizeHttpCredentials(httpCredentials) {
  if (!httpCredentials) return undefined;

  const username = String(httpCredentials.username || "").trim();
  const password = String(httpCredentials.password || "");

  if (!username && !password) return undefined;
  if (!username || !password) {
    throw new Error("HTTP Basic Auth must include both username and password.");
  }

  return { username, password };
}

function normalizeBrowserEngine(browserEngine) {
  const value = String(browserEngine || "chromium").trim().toLowerCase();
  if (value === "chromium" || value === "chrome") return value;
  throw new Error(`Unsupported browser engine: ${browserEngine}`);
}

function resolveChromeCdpUrl(options = {}) {
  const explicit = String(options.chromeCdpUrl || "").trim();
  if (explicit) return explicit;

  const envUrl = String(process.env.CHROME_CDP_URL || "").trim();
  if (envUrl) return envUrl;

  const debugPort = String(process.env.CHROME_DEBUG_PORT || "").trim();
  if (debugPort) return `http://127.0.0.1:${debugPort}`;

  return DEFAULT_CHROME_CDP_URL;
}

async function withScriptedClose(task) {
  _closingByScript = true;
  try {
    return await task();
  } finally {
    await new Promise((resolve) => setTimeout(resolve, 0));
    _closingByScript = false;
  }
}

function markSessionClosed(session, reason) {
  if (!session) return;
  session.closed = true;
  if (!session.closeReason) {
    session.closeReason = reason;
  }
}

function handleUnexpectedSessionClose(session, reason) {
  markSessionClosed(session, reason);

  if (_closingByScript || _manualCloseExitPending) {
    return;
  }

  _manualCloseExitPending = true;
  _session = null;
  process.stdin.pause();
  console.log(`[browserSession] ${reason}. Exiting.`);
  process.exit(0);
}

function attachUnexpectedCloseHandlers({ session, context, page, browser }) {
  page.on("close", () => {
    handleUnexpectedSessionClose(session, "Browser page was closed");
  });

  context.on("close", () => {
    handleUnexpectedSessionClose(session, "Browser window was closed");
  });

  if (browser && typeof browser.on === "function") {
    browser.on("disconnected", () => {
      handleUnexpectedSessionClose(session, "Browser connection was closed");
    });
  }
}

function normalizeLaunchOptions(options = {}) {
  const attachToExistingChrome = Boolean(options.attachToExistingChrome);
  const browserEngine = attachToExistingChrome
    ? "chrome"
    : normalizeBrowserEngine(options.browserEngine);

  return {
    headless: Boolean(options.headless),
    window: resolveWindowOptions(options.window),
    httpCredentials: normalizeHttpCredentials(options.httpCredentials),
    browserEngine,
    attachToExistingChrome,
    chromeCdpUrl: resolveChromeCdpUrl(options),
  };
}

function mergeLaunchOptions(baseOptions = {}, patchOptions = {}) {
  const merged = {
    ...baseOptions,
    ...patchOptions,
    window: {
      ...(baseOptions.window || {}),
      ...(patchOptions.window || {}),
    },
    httpCredentials:
      patchOptions.httpCredentials === undefined
        ? baseOptions.httpCredentials
        : patchOptions.httpCredentials,
    attachToExistingChrome:
      patchOptions.attachToExistingChrome === undefined
        ? baseOptions.attachToExistingChrome
        : patchOptions.attachToExistingChrome,
    browserEngine:
      patchOptions.browserEngine === undefined
        ? baseOptions.browserEngine
        : patchOptions.browserEngine,
    chromeCdpUrl:
      patchOptions.chromeCdpUrl === undefined
        ? baseOptions.chromeCdpUrl
        : patchOptions.chromeCdpUrl,
  };

  return normalizeLaunchOptions(merged);
}

function registerSignalHandlers() {
  if (_signalsRegistered) return;

  process.once("SIGINT", async () => {
    await cleanupBrowserSession();
    process.exit(130);
  });

  process.once("SIGTERM", async () => {
    await cleanupBrowserSession();
    process.exit(143);
  });

  process.once("exit", () => {
    if (_lockAcquired) releaseLock();
  });

  _signalsRegistered = true;
}

async function launchPersistentContext(launchOptions) {
  const { window, headless, httpCredentials, browserEngine } = launchOptions;

  console.log(`[browserSession] Using persistent profile: ${PROFILE_DIR}`);
  console.log(
    `[browserSession] Window: ${window.width}x${window.height} @ (${window.positionX},${window.positionY})`
  );
  console.log(`[browserSession] Engine: ${browserEngine}`);
  console.log(`[browserSession] Mode: ${headless ? "headless" : "headed"}`);
  if (httpCredentials?.username) {
    console.log("[browserSession] HTTP Basic Auth: configured");
  } else {
    console.log("[browserSession] HTTP Basic Auth: not configured");
  }

  const launchArgs = {
    headless,
    viewport: { width: window.width, height: window.height },
    httpCredentials,
    args: [
      `--window-size=${window.width},${window.height}`,
      `--window-position=${window.positionX},${window.positionY}`,
      "--disable-blink-features=AutomationControlled",
      "--no-first-run",
      "--no-default-browser-check",
      "--hide-crash-restore-bubble",
      "--disable-session-crashed-bubble",
    ],
  };

  if (browserEngine === "chrome") {
    launchArgs.channel = "chrome";
  }

  const context = await chromium.launchPersistentContext(PROFILE_DIR, launchArgs);

  const page = context.pages()[0] || (await context.newPage());
  if (!headless) {
    await page.bringToFront().catch(() => {});
  }

  _lastLaunchOptions = launchOptions;
  _manualCloseExitPending = false;
  const session = {
    context,
    page,
    profileDir: PROFILE_DIR,
    cleanup: cleanupBrowserSession,
    headless,
    browserEngine,
    isExternalBrowser: false,
    closed: false,
    closeReason: null,
    isClosed() {
      try {
        return Boolean(this.closed) || page.isClosed();
      } catch (_) {
        return Boolean(this.closed);
      }
    },
    safeClose: async () => {
      markSessionClosed(session, "Closed by automation");
      await withScriptedClose(async () => {
        try {
          await context.close();
        } catch (_) {
          // Ignore close errors during cleanup/relaunch.
        }
      });
    },
  };
  _session = session;
  attachUnexpectedCloseHandlers({ session, context, page, browser: context.browser() });

  return session;
}

async function connectExistingChromeSession(launchOptions) {
  const { chromeCdpUrl, httpCredentials } = launchOptions;

  console.log("[browserSession] Attaching to existing Chrome via CDP");
  console.log(`[browserSession] CDP URL: ${chromeCdpUrl}`);
  if (httpCredentials?.username) {
    console.log(
      "[browserSession] HTTP Basic Auth is ignored when attaching to existing Chrome context."
    );
  }

  const browser = await chromium.connectOverCDP(chromeCdpUrl, {
    timeout: 15000,
  });
  const context = browser.contexts()[0];

  if (!context) {
    throw new Error(
      "Connected to Chrome CDP, but no browser context was available. Keep Chrome open and retry."
    );
  }

  const page = await context.newPage();
  await page.bringToFront().catch(() => {});

  _lastLaunchOptions = launchOptions;
  _manualCloseExitPending = false;
  const session = {
    context,
    page,
    profileDir: null,
    cleanup: cleanupBrowserSession,
    headless: false,
    browserEngine: "chrome",
    isExternalBrowser: true,
    closed: false,
    closeReason: null,
    isClosed() {
      try {
        return Boolean(this.closed) || page.isClosed();
      } catch (_) {
        return Boolean(this.closed);
      }
    },
    safeClose: async () => {
      markSessionClosed(session, "Closed by automation");
      await withScriptedClose(async () => {
        try {
          if (!page.isClosed()) {
            await page.close();
          }
        } catch (_) {
          // Ignore page close errors on external browser sessions.
        }
        try {
          await browser.close();
        } catch (_) {
          // Ignore browser close/disconnect errors for CDP sessions.
        }
      });
    },
  };
  _session = session;
  attachUnexpectedCloseHandlers({ session, context, page, browser });

  return session;
}

function ensureBrowserProfileLock() {
  if (_lockAcquired) return;
  acquireLock();
  markChromiumProfileAsClean();
}

/**
 * @param {object} [options]
 * @param {boolean} [options.headless]
 * @param {{username:string,password:string}} [options.httpCredentials] - HTTP Basic Auth credentials
 * @param {{width:number,height:number,positionX:number,positionY:number}} [options.window]
 */
async function getBrowserSession(options = {}) {
  if (_session) return _session;
  if (_sessionPromise) return _sessionPromise;

  _sessionPromise = (async () => {
    registerSignalHandlers();
    const launchOptions = normalizeLaunchOptions(options);
    if (!launchOptions.attachToExistingChrome) {
      ensureBrowserProfileLock();
      return launchPersistentContext(launchOptions);
    }
    return connectExistingChromeSession(launchOptions);
  })().finally(() => {
    _sessionPromise = null;
  });

  return _sessionPromise;
}

/**
 * Relaunches the persistent browser context with updated options
 * (e.g. switch from headless -> headed) while keeping the profile lock.
 */
async function relaunchBrowserSession(options = {}) {
  if (!_session) {
    return getBrowserSession(options);
  }

  const launchOptions = mergeLaunchOptions(_lastLaunchOptions || {}, options);
  const currentSession = _session;
  _session = null;

  await currentSession.safeClose();

  if (!launchOptions.attachToExistingChrome) {
    ensureBrowserProfileLock();
    markChromiumProfileAsClean();
    return launchPersistentContext(launchOptions);
  }

  return connectExistingChromeSession(launchOptions);
}

async function cleanupBrowserSession() {
  const current = _session;
  _session = null;

  if (current?.safeClose) {
    await current.safeClose();
  } else if (current?.context) {
    markSessionClosed(current, "Closed by automation");
    await withScriptedClose(async () => {
      try {
        await current.context.close();
      } catch (_) {}
    });
  }

  if (_lockAcquired) {
    releaseLock();
  }
}

module.exports = {
  getBrowserSession,
  relaunchBrowserSession,
  PROFILE_DIR,
};
