// automation/betterLetterBasicAuth.js
const fs = require("fs");
const path = require("path");

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
const AUTH_FILE = path.join(PROJECT_ROOT, ".betterletter-basic-auth.json");

function loadBetterLetterBasicAuth() {
  // Prefer environment variables (best practice for secrets)
  const envUser = process.env.BETTERLETTER_BASIC_AUTH_USER;
  const envPass = process.env.BETTERLETTER_BASIC_AUTH_PASS;

  if (envUser && envPass) {
    return { username: envUser, password: envPass };
  }

  // Fallback to local file (M: only on your machine
  if (fs.existsSync(AUTH_FILE)) {
    try {
      const raw = fs.readFileSync(AUTH_FILE, "utf8");
      const parsed = JSON.parse(raw);
      if (parsed?.username && parsed?.password) {
        return { username: parsed.username, password: parsed.password };
      }
    } catch (_) {
      // ignore and behave as "not set"
    }
  }

  return null;
}

function saveBetterLetterBasicAuth({ username, password }) {
  if (!username || !password) throw new Error("Cannot save empty credentials.");

  fs.writeFileSync(
    AUTH_FILE,
    JSON.stringify({ username, password }, null, 2),
    "utf8"
  );

  // best-effort: restrict permissions on mac/linux
  try {
    fs.chmodSync(AUTH_FILE, 0o600);
  } catch (_) {}

  return AUTH_FILE;
}

module.exports = {
  loadBetterLetterBasicAuth,
  saveBetterLetterBasicAuth,
  AUTH_FILE,
};
