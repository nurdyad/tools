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
const DEFAULT_ENV_FILE = path.join(PROJECT_ROOT, ".env");

function parseEnvValue(raw) {
  const value = String(raw || "").trim();
  if (!value) return "";

  const quote = value[0];
  if ((quote === '"' || quote === "'") && value[value.length - 1] === quote) {
    const inner = value.slice(1, -1);
    return inner
      .replace(/\\n/g, "\n")
      .replace(/\\r/g, "\r")
      .replace(/\\t/g, "\t");
  }

  return value;
}

function loadEnvKeyValues(filePath) {
  if (!fs.existsSync(filePath)) {
    return { filePath, loadedCount: 0, exists: false };
  }

  const content = fs.readFileSync(filePath, "utf8");
  const lines = content.split(/\r?\n/);
  let loadedCount = 0;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;

    const match = trimmed.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
    if (!match) continue;

    const key = match[1];
    const rawValue = match[2] || "";
    if (process.env[key] != null && process.env[key] !== "") continue;

    process.env[key] = parseEnvValue(rawValue);
    loadedCount += 1;
  }

  return { filePath, loadedCount, exists: true };
}

function loadEnvFile(filePath = DEFAULT_ENV_FILE) {
  return loadEnvKeyValues(filePath);
}

module.exports = {
  DEFAULT_ENV_FILE,
  loadEnvFile,
};
