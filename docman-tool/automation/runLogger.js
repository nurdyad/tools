const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

function classifyError(error) {
  if (!error) return "UnknownError";
  if (typeof error === "string") return "Error";
  if (error.name && typeof error.name === "string") return error.name;
  if (error.constructor?.name) return error.constructor.name;
  return "Error";
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function createRunLogger(options = {}) {
  const {
    enabled = true,
    projectRoot = process.cwd(),
    logDirectory = "logs",
    app = "docman-tool",
  } = options;

  const sessionId = crypto.randomBytes(8).toString("hex");

  if (!enabled) {
    return {
      enabled: false,
      sessionId,
      filePath: null,
      event: () => {},
      startStep: () => ({ name: "disabled", startedAt: Date.now() }),
      endStep: () => {},
      close: () => {},
    };
  }

  const dir = path.join(projectRoot, logDirectory);
  ensureDir(dir);

  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const filePath = path.join(dir, `${app}-run-${stamp}-${sessionId}.jsonl`);
  const stream = fs.createWriteStream(filePath, { flags: "a" });

  let streamHealthy = true;
  let streamErrorNotified = false;

  function markStreamFailed(error) {
    streamHealthy = false;
    if (streamErrorNotified) return;
    streamErrorNotified = true;

    const message = error?.message || String(error || "unknown stream error");
    console.warn(`Log writing disabled: ${message}`);
  }

  stream.on("error", (error) => {
    markStreamFailed(error);
  });

  function writeLine(payload) {
    if (!streamHealthy) return;

    const line = JSON.stringify({
      timestamp: new Date().toISOString(),
      sessionId,
      ...payload,
    });

    try {
      stream.write(`${line}\n`);
    } catch (error) {
      markStreamFailed(error);
    }
  }

  function event(eventName, data = {}) {
    writeLine({ event: eventName, ...data });
  }

  function startStep(name, meta = {}) {
    const startedAt = Date.now();
    event("step_start", { step: name, ...meta });
    return { name, startedAt };
  }

  function endStep(step, data = {}) {
    const durationMs = Math.max(0, Date.now() - (step?.startedAt || Date.now()));
    event("step_end", {
      step: step?.name || "unknown",
      durationMs,
      ...data,
    });
  }

  function close(data = {}) {
    if (streamHealthy) {
      event("run_end", data);
    }

    try {
      stream.end();
    } catch (_) {}
  }

  event("run_start", { app, cwd: process.cwd(), pid: process.pid });

  return {
    enabled: true,
    sessionId,
    filePath,
    event,
    startStep,
    endStep,
    close,
  };
}

module.exports = {
  classifyError,
  createRunLogger,
};
