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

/**
 * @param {object} [options]
 * @param {{username:string,password:string}} [options.httpCredentials] - HTTP Basic Auth credentials
 */
async function getBrowserSession(options = {}) {
  if (_sessionPromise) return _sessionPromise;

  _sessionPromise = (async () => {
    acquireLock();

    console.log(`[browserSession] Using persistent profile: ${PROFILE_DIR}`);
    if (options.httpCredentials?.username) {
      console.log("[browserSession] HTTP Basic Auth: configured");
    } else {
      console.log("[browserSession] HTTP Basic Auth: not configured");
    }

    const context = await chromium.launchPersistentContext(PROFILE_DIR, {
      headless: false,
      viewport: null,
      httpCredentials: options.httpCredentials, // ✅ this is the fix for the popup

      args: [
        "--start-maximized",
        "--disable-blink-features=AutomationControlled",
        "--no-first-run",
        "--no-default-browser-check",
      ],
    });

    const page = context.pages()[0] || (await context.newPage());
    await page.bringToFront();

    const cleanup = async () => {
      try {
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
