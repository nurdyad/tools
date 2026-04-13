function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function withRetry(task, options = {}) {
  const {
    attempts = 1,
    baseDelayMs = 0,
    maxDelayMs = baseDelayMs,
    label = "operation",
    onRetry,
  } = options;

  const safeAttempts = Number.isInteger(attempts) && attempts > 0 ? attempts : 1;
  const safeBaseDelay = Math.max(0, Number(baseDelayMs) || 0);
  const safeMaxDelay = Math.max(safeBaseDelay, Number(maxDelayMs) || safeBaseDelay);

  let lastError = null;

  for (let attempt = 1; attempt <= safeAttempts; attempt++) {
    try {
      return await task({ attempt, attempts: safeAttempts, label });
    } catch (error) {
      lastError = error;
      if (attempt >= safeAttempts) break;

      const nextAttempt = attempt + 1;
      const rawDelay = safeBaseDelay * Math.pow(2, attempt - 1);
      const delayMs = Math.min(safeMaxDelay, rawDelay);

      if (typeof onRetry === "function") {
        onRetry({
          label,
          attempt,
          attempts: safeAttempts,
          nextAttempt,
          delayMs,
          error,
        });
      }

      if (delayMs > 0) {
        await sleep(delayMs);
      }
    }
  }

  throw lastError || new Error(`${label} failed after ${safeAttempts} attempt(s)`);
}

module.exports = {
  withRetry,
};
