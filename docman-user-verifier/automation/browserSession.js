// automation/browserSession.js
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");

function findProjectRoot(startDir) {
  let dir = startDir;
  while (true) {
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
const WINDOW_WIDTH = 1920;
const WINDOW_HEIGHT = 1080;

let _sessionPromise = null;

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function acquireLock() {
  ensureDir(PROFILE_DIR);

  if (fs.existsSync(LOCK_FILE)) {
    const lockContent = fs.readFileSync(LOCK_FILE, "utf8").trim();
    throw new Error(
      [
        `Profile appears to be in use (lock file exists): ${LOCK_FILE}`,
        lockContent ? `Lock info: ${lockContent}` : "",
        "Close any other running instance of this automation (or Chromium using this profile) and try again.",
      ].filter(Boolean).join("\n")
    );
  }

  fs.writeFileSync(
    LOCK_FILE,
    `pid=${process.pid}\nstarted=${new Date().toISOString()}\nprofile=${PROFILE_DIR}\n`,
    "utf8"
  );
}

function releaseLock() {
  try {
    if (fs.existsSync(LOCK_FILE)) fs.unlinkSync(LOCK_FILE);
  } catch (_) {}
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

/**
 * @param {object} [options]
 * @param {{username:string,password:string}} [options.httpCredentials] - HTTP Basic Auth credentials
 */
async function getBrowserSession(options = {}) {
  if (_sessionPromise) return _sessionPromise;

  _sessionPromise = (async () => {
    acquireLock();
    markChromiumProfileAsClean();

    console.log(`[browserSession] Using persistent profile: ${PROFILE_DIR}`);
    if (options.httpCredentials?.username) {
      console.log("[browserSession] HTTP Basic Auth: configured");
    } else {
      console.log("[browserSession] HTTP Basic Auth: not configured");
    }

    const context = await chromium.launchPersistentContext(PROFILE_DIR, {
      headless: false,
      viewport: { width: WINDOW_WIDTH, height: WINDOW_HEIGHT },
      httpCredentials: options.httpCredentials, // ✅ this is the fix for the popup

      args: [
        `--window-size=${WINDOW_WIDTH},${WINDOW_HEIGHT}`,
        "--window-position=40,40",
        "--disable-blink-features=AutomationControlled",
        "--no-first-run",
        "--no-default-browser-check",
        "--hide-crash-restore-bubble",
        "--disable-session-crashed-bubble",
      ],
    });

    const page = context.pages()[0] || (await context.newPage());
    await page.bringToFront();

    // Exit Node when the browser window is closed

    let closingByScript = false;

    context.on("close", () => {
      if (!closingByScript) {
        console.log("[browserSession] Browser window closed — exiting.");
        process.exit(0);
      }
    });

    // When you close it yourself (normal shutdown), set the flag:
    const cleanup = async () => {
      try {
        closingByScript = true;
        await context.close();
      } catch (_) {}
      releaseLock();
    };

    process.once("SIGINT", async () => {
      await cleanup();
      process.exit(130);
    });

    process.once("SIGTERM", async () => {
      await cleanup();
      process.exit(143);
    });

    process.once("exit", () => {
      releaseLock();
    });

    return { context, page, profileDir: PROFILE_DIR, cleanup };
  })();

  return _sessionPromise;
}

module.exports = { getBrowserSession, PROFILE_DIR };
