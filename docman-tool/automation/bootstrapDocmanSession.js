// automation/bootstrapDocmanSession.js
const { getBrowserSession, relaunchBrowserSession } = require("./browserSession");
const { fetchDocmanCreds } = require("./fetchDocmanCreds");
const { withRetry } = require("./retry");
const { classifyError } = require("./runLogger");

const BETTERLETTER_PRACTICES_URL = "https://app.betterletter.ai/admin_panel/practices";
const DOCMAN_ORIGIN = "https://production.docman.thirdparty.nhs.uk";
const DOCMAN_HOST_SUFFIX = "docman.thirdparty.nhs.uk";
const DOCMAN_FILING_URL = `${DOCMAN_ORIGIN}/DocumentViewer/Filing`;
const DOCMAN_LOGOUT_URL = `${DOCMAN_ORIGIN}/Account/Logout`;
const DOCMAN_LOGIN_URL = `${DOCMAN_ORIGIN}/Account/Login`;

async function bootstrapDocmanSession(practiceInput, sessionOptions = {}) {
  const {
    skipPostLoginDialogWatch = false,
    forceFreshDocmanLogin = false,
    resetDocmanAuthAtStart = true,
    includeDocmanInHealthCheck = false,
    preDocmanHeadless = true,
    allowVisibleFallback = false,
    step4Visible = true,
    step4UseCurrentChrome = true,
    step4BrowserEngine = "chrome",
    chromeCdpUrl = "",
    retryPolicy = {},
    timeouts = {},
    logger = null,
  } = sessionOptions;

  const resolvedTimeouts = normalizeTimeouts(timeouts);
  const resolvedRetryPolicy = normalizeRetryPolicy(retryPolicy);
  const preferredStep4Engine =
    String(step4BrowserEngine || "chrome").trim().toLowerCase() === "chromium"
      ? "chromium"
      : "chrome";
  let runningHeadless = Boolean(preDocmanHeadless);
  let browserSession = await getBrowserSession({
    ...sessionOptions,
    headless: runningHeadless,
  });
  let { context, page } = browserSession;

  // Always start on BetterLetter admin panel when the browser opens.
  await page
    .goto(BETTERLETTER_PRACTICES_URL, {
      waitUntil: "domcontentloaded",
      timeout: resolvedTimeouts.navigationMs,
    })
    .catch(() => {});

  const planLines = [
    "1) Session health check",
    runningHeadless
      ? "2) Login to BetterLetter (automatic sign-in + automatic 2FA when configured)"
      : "2) Login to BetterLetter (manual fallback available)",
    "3) Read Docman credentials from BetterLetter",
    step4Visible
      ? step4UseCurrentChrome
        ? "4) Login to Docman in current Chrome window (visible)"
        : "4) Open visible Chrome and login to Docman"
      : "4) Login to Docman with BetterLetter credentials",
  ];
  if (!skipPostLoginDialogWatch) {
    planLines.push("5) Dismiss blocking dialogs");
  }

  printPhaseBanner("DOCMAN BOOTSTRAP", planLines);

  // Practice is collected in run.js before browser launch.
  const practiceName =
    typeof practiceInput === "function"
      ? await practiceInput({ page })
      : practiceInput;

  if (!practiceName || !practiceName.trim()) {
    throw new Error("Practice name is required.");
  }

  if (resetDocmanAuthAtStart) {
    const removed = await clearDocmanAuthFromContext(context);
    console.log(`🧹 Cleared Docman auth artifacts at startup (${removed} cookie(s) removed).`);
  }

  await runBootstrapStep({
    stepNumber: 1,
    title: "Session health check",
    task: () =>
      sessionHealthCheck(page, {
        includeDocmanCheck: includeDocmanInHealthCheck,
        timeouts: resolvedTimeouts,
      }),
    retryPolicy: resolvedRetryPolicy,
    logger,
  });

  try {
    await runBootstrapStep({
      stepNumber: 2,
      title: "Login to BetterLetter",
      task: () =>
        ensureBetterLetterLoggedIn(page, {
          timeouts: resolvedTimeouts,
          allowManualInput: !runningHeadless,
        }),
      retryPolicy: runningHeadless
        ? { ...resolvedRetryPolicy, attempts: 1 }
        : resolvedRetryPolicy,
      logger,
    });
  } catch (error) {
    if (!runningHeadless || !isInteractiveAuthRequiredError(error) || !allowVisibleFallback) {
      throw error;
    }

    console.log("🖥 BetterLetter requires interactive input. Opening visible Chrome and retrying login...");
    browserSession = await relaunchBrowserSession({
      ...sessionOptions,
      headless: false,
    });
    ({ context, page } = browserSession);
    runningHeadless = false;

    await runBootstrapStep({
      stepNumber: 2,
      title: "Login to BetterLetter",
      task: () =>
        ensureBetterLetterLoggedIn(page, {
          timeouts: resolvedTimeouts,
          allowManualInput: true,
        }),
      retryPolicy: resolvedRetryPolicy,
      logger,
    });
  }

  let creds;
  try {
    creds = await runBootstrapStep({
      stepNumber: 3,
      title: "Read Docman ODS/username/password from BetterLetter",
      task: () => fetchDocmanCreds(page, practiceName.trim()),
      retryPolicy: resolvedRetryPolicy,
      logger,
    });
  } catch (error) {
    if (!runningHeadless || !allowVisibleFallback) {
      throw error;
    }

    console.log(
      `🖥 Headless credential fetch failed (${error?.message || "unknown error"}). ` +
      "Opening visible Chrome and retrying Step 3..."
    );

    browserSession = await relaunchBrowserSession({
      ...sessionOptions,
      headless: false,
    });
    ({ context, page } = browserSession);
    runningHeadless = false;

    creds = await runBootstrapStep({
      stepNumber: 3,
      title: "Read Docman ODS/username/password from BetterLetter",
      task: () => fetchDocmanCreds(page, practiceName.trim()),
      retryPolicy: resolvedRetryPolicy,
      logger,
    });
  }

  const {
    odsCode,
    adminUsername,
    adminPassword,
    inputFolder,
    processingFolder,
    filingFolder,
    rejectedFolder,
  } = creds;

  if (step4Visible) {
    const attachToExistingChrome = Boolean(step4UseCurrentChrome);
    const requiresStep4Relaunch =
      runningHeadless ||
      attachToExistingChrome ||
      String(browserSession?.browserEngine || "").toLowerCase() !== preferredStep4Engine;

    if (requiresStep4Relaunch) {
      console.log(
        attachToExistingChrome
          ? "🖥 Switching Step 4 to current Chrome window..."
          : "🖥 Switching Step 4 to visible Chrome..."
      );

      try {
        browserSession = await relaunchBrowserSession({
          ...sessionOptions,
          headless: false,
          browserEngine: preferredStep4Engine,
          attachToExistingChrome,
          chromeCdpUrl,
        });
        ({ context, page } = browserSession);
      } catch (error) {
        if (attachToExistingChrome && isCdpAttachConnectionError(error)) {
          console.log(
            "⚠ Could not attach to current Chrome (CDP unavailable). " +
              "Falling back to a visible Chrome window for Step 4."
          );
          console.log(
            "ℹ To use the current Chrome window next run, start Chrome with: " +
              "\"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome --remote-debugging-port=9222\""
          );

          browserSession = await relaunchBrowserSession({
            ...sessionOptions,
            headless: false,
            browserEngine: "chrome",
            attachToExistingChrome: false,
          });
          ({ context, page } = browserSession);
        } else if (attachToExistingChrome) {
          throw new Error(
            "Could not attach to the current Chrome window for Step 4. " +
              "Start Chrome with remote debugging first, e.g. " +
              "\"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome --remote-debugging-port=9222\". " +
              `Details: ${error?.message || String(error)}`
          );
        }
        else {
          throw error;
        }
      }

      runningHeadless = false;

      if (resetDocmanAuthAtStart) {
        const removed = await clearDocmanAuthFromContext(context);
        console.log(`🧹 Cleared Docman auth artifacts before Step 4 (${removed} cookie(s) removed).`);
      }
    }
  }

  await runBootstrapStep({
    stepNumber: 4,
    title: "Login to Docman with BetterLetter credentials",
    task: () =>
      ensureDocmanSessionForPractice(
        page,
        {
          practiceName: practiceName.trim(),
          odsCode,
          adminUsername,
          adminPassword,
        },
        {
          forceFreshLogin: forceFreshDocmanLogin,
          skipExistingSessionCheck: resetDocmanAuthAtStart,
          timeouts: resolvedTimeouts,
          allowManualFallback: !runningHeadless,
        }
      ),
    retryPolicy: resolvedRetryPolicy,
    logger,
  });

  // 5) Clear any post-login modal(s) when requested by workflow.
  if (!skipPostLoginDialogWatch) {
    await runBootstrapStep({
      stepNumber: 5,
      title: "Dismiss blocking dialogs",
      task: () => waitAndDismissBlockingDialogs(page, "after docman login"),
      retryPolicy: resolvedRetryPolicy,
      logger,
    });
  }

  return {
    context,
    page,
    odsCode,
    adminUsername,
    adminPassword,
    inputFolder,
    processingFolder,
    filingFolder,
    rejectedFolder,
    headless: runningHeadless,
    browserEngine: browserSession?.browserEngine,
    isExternalBrowser: Boolean(browserSession?.isExternalBrowser),
    safeClose: browserSession?.safeClose,
  };
}

function normalizeTimeouts(timeouts = {}) {
  const read = (key, fallback) => {
    const value = Number(timeouts?.[key]);
    return Number.isFinite(value) && value > 0 ? Math.floor(value) : fallback;
  };

  return {
    navigationMs: read("navigationMs", 60000),
    selectorMs: read("selectorMs", 60000),
    docmanLoginFastCheckMs: read("docmanLoginFastCheckMs", 3500),
    docmanLoginDeepCheckMs: read("docmanLoginDeepCheckMs", 9000),
  };
}

function normalizeRetryPolicy(retryPolicy = {}) {
  const attempts = Number.isInteger(retryPolicy?.attempts) ? retryPolicy.attempts : 1;
  const baseDelayMs = Number.isFinite(retryPolicy?.baseDelayMs)
    ? Math.max(0, Math.floor(retryPolicy.baseDelayMs))
    : 0;
  const maxDelayMs = Number.isFinite(retryPolicy?.maxDelayMs)
    ? Math.max(baseDelayMs, Math.floor(retryPolicy.maxDelayMs))
    : baseDelayMs;

  return {
    attempts: Math.max(1, attempts),
    baseDelayMs,
    maxDelayMs,
  };
}

async function runBootstrapStep({ stepNumber, title, task, retryPolicy, logger }) {
  printPhaseBanner(`Step ${stepNumber}`, [title]);

  const stepKey = `bootstrap_step_${stepNumber}`;
  const stepToken = logger?.startStep(stepKey, { title });

  try {
    const result = await withRetry(task, {
      ...retryPolicy,
      label: title,
      onRetry: ({ nextAttempt, attempts, delayMs, error }) => {
        const errorMessage = error?.message || "unknown error";
        console.log(
          `↻ Step ${stepNumber} failed (${errorMessage}). retry ${nextAttempt}/${attempts} in ${delayMs}ms`
        );
        logger?.event("bootstrap_step_retry", {
          step: stepKey,
          title,
          nextAttempt,
          attempts,
          delayMs,
          errorType: classifyError(error),
          errorMessage,
        });
      },
    });

    logger?.endStep(stepToken, { status: "ok" });
    return result;
  } catch (error) {
    logger?.endStep(stepToken, {
      status: "error",
      errorType: classifyError(error),
      errorMessage: error?.message || String(error),
    });
    throw error;
  }
}
/* ------------ session health check ------------ */

async function sessionHealthCheck(page, options = {}) {
  console.log("\n================ SESSION HEALTH CHECK ================");
  const includeDocmanCheck = Boolean(options.includeDocmanCheck);
  const navigationTimeoutMs = Number(options?.timeouts?.navigationMs) || 45000;
  const authCheckTimeoutMs = Number(options?.timeouts?.docmanLoginFastCheckMs) || 6000;

  // ---------------- BetterLetter check ----------------
  let blLoggedIn = false;
  try {
    const resp = await page.goto("https://app.betterletter.ai/admin_panel/practices", {
      waitUntil: "domcontentloaded",
      timeout: navigationTimeoutMs,
    });

    const status = resp?.status?.();
    const u = page.url();

    const bodyText = (await page.textContent("body").catch(() => "")) || "";
    const unauthorized =
      status === 401 ||
      status === 403 ||
      /unauthorized/i.test(bodyText);

    const onSignIn =
      u.includes("/sign-in") ||
      u.includes("/users/log_in") ||
      u.includes("/login");

    blLoggedIn = !unauthorized && !onSignIn;

    console.log(
      `BetterLetter debug: status=${status ?? "n/a"} unauthorized=${unauthorized} url=${u}`
    );
  } catch (e) {
    console.log("BetterLetter: ⚠ Unable to determine (navigation issue)");
  }

  console.log(`BetterLetter: ${blLoggedIn ? "✅ Logged in" : "❌ Not logged in"}`);


  if (includeDocmanCheck) {
    // ---------------- Docman check ----------------
    let dmLoggedIn = false;
    try {
      const resp = await page.goto(DOCMAN_FILING_URL, {
        waitUntil: "domcontentloaded",
        timeout: navigationTimeoutMs,
      });

      // Give redirects a moment
      await page.waitForTimeout(300);

      const authSurface = await inspectDocmanAuthSurface(page, authCheckTimeoutMs);
      const onLoginPage = authSurface.onLoginPage;

      dmLoggedIn = !onLoginPage;

      console.log(
        `Docman debug: status=${resp?.status?.() ?? "n/a"} url=${authSurface.url} onLoginPage=${onLoginPage}`
      );
    } catch (e) {
      console.log("Docman: ⚠ Unable to determine (navigation issue)");
    }

    console.log(`Docman: ${dmLoggedIn ? "✅ Logged in" : "❌ Not logged in"}`);
  } else {
    console.log("Docman: ⏭ Skipped in health check (keeping BetterLetter page focused).");
  }

  console.log("======================================================\n");

  // Keep BetterLetter page in focus after health-check.
  try {
    await page.goto(BETTERLETTER_PRACTICES_URL, {
      waitUntil: "domcontentloaded",
      timeout: navigationTimeoutMs,
    });
  } catch (_) {}
}


/* ------------ BetterLetter helpers ------------ */

async function ensureBetterLetterLoggedIn(page, options = {}) {
  const navigationTimeoutMs = Number(options?.timeouts?.navigationMs) || 60000;
  const selectorTimeoutMs = Number(options?.timeouts?.selectorMs) || 60000;
  const authWaitTimeoutMs = Number(process.env.POST_LOGIN_WAIT_SECONDS || 120) * 1000;
  const allowManualInput = options?.allowManualInput !== false;

  await page.goto(BETTERLETTER_PRACTICES_URL, {
    waitUntil: "domcontentloaded",
    timeout: navigationTimeoutMs,
  });

  // Unauthorized page body check (basic auth failure / access denied)
  const bodyText = ((await page.textContent("body").catch(() => "")) || "").trim();
  if (/unauthorized/i.test(bodyText)) {
    throw new Error(
      "BetterLetter returned Unauthorized (HTTP Basic Auth failed). Check BETTERLETTER_BASIC_AUTH_USER/BETTERLETTER_BASIC_AUTH_PASS or .betterletter-basic-auth.json."
    );
  }

  let authSurface = await inspectBetterLetterAuthSurface(page);

  if (authSurface.onLoginPage) {
    const credentials = getBetterLetterPrimaryCredentials();
    let autoPrimarySubmitted = false;

    if (credentials) {
      console.log("🔐 BetterLetter sign-in detected. Attempting auto sign-in...");
      autoPrimarySubmitted = await fillBetterLetterPrimaryLogin(
        page,
        credentials.email,
        credentials.password
      );
    }

    if (autoPrimarySubmitted) {
      try {
        await waitForBetterLetter2faOrPractices(page, authWaitTimeoutMs);
      } catch (error) {
        const retrySurface = await inspectBetterLetterAuthSurface(page);
        if (!retrySurface.onLoginPage) throw error;

        console.log("↻ BetterLetter sign-in did not advance. Retrying once...");
        await fillBetterLetterPrimaryLogin(page, credentials.email, credentials.password);
        await waitForBetterLetter2faOrPractices(page, authWaitTimeoutMs);
      }
    }

    authSurface = await inspectBetterLetterAuthSurface(page);
  }

  if (authSurface.on2faPage) {
    const auto2faFromEmail = shouldAutoFetch2faFromEmail();
    await completeBetterLetter2fa(page, {
      auto2faFromEmail,
      authWaitTimeoutMs,
      allowManualInput,
    });
    authSurface = await inspectBetterLetterAuthSurface(page);
  }

  if (authSurface.onLoginPage || authSurface.on2faPage) {
    if (!allowManualInput) {
      throw createInteractiveAuthRequiredError(
        authSurface.on2faPage
          ? "BetterLetter needs manual 2FA input, but browser is headless."
          : "BetterLetter needs manual sign-in input, but browser is headless."
      );
    }

    const promptText = authSurface.on2faPage
      ? "🔐 BetterLetter requires the 2FA security code. Enter it in the browser and press ENTER here."
      : "🔐 Please log into BetterLetter in the opened browser window, then press ENTER here.";
    console.log(`\n${promptText}\n`);
    await waitForEnter();
  }

  // Ensure we're back on practices after any sign-in / 2FA hand-off.
  await page.goto(BETTERLETTER_PRACTICES_URL, {
    waitUntil: "domcontentloaded",
    timeout: navigationTimeoutMs,
  });

  await page.waitForSelector('a[href^="/admin_panel/practices/"]', {
    timeout: selectorTimeoutMs,
  });

  if (isBetterLetterSignInUrl(page.url())) {
    throw new Error("BetterLetter still shows sign-in after login attempt.");
  }

  return true;
}

function envFlag(value, fallback = false) {
  if (value == null || String(value).trim() === "") return fallback;
  return ["1", "true", "yes", "on"].includes(String(value).trim().toLowerCase());
}

function parseBooleanEnv(value) {
  const normalized = envValue(value).toLowerCase();
  if (!normalized) return null;
  if (["1", "true", "yes", "on"].includes(normalized)) return true;
  if (["0", "false", "no", "off"].includes(normalized)) return false;
  return null;
}

function envValue(value) {
  return String(value ?? "").replace(/\r/g, "").replace(/\n/g, "").trim();
}

function createInteractiveAuthRequiredError(message) {
  const error = new Error(message || "Interactive BetterLetter input is required.");
  error.code = "BETTERLETTER_INTERACTIVE_REQUIRED";
  return error;
}

function isInteractiveAuthRequiredError(error) {
  return String(error?.code || "").toUpperCase() === "BETTERLETTER_INTERACTIVE_REQUIRED";
}

function shouldAutoFetch2faFromEmail() {
  const explicitToggle = envValue(process.env.AUTO_2FA_FROM_EMAIL);
  const parsedToggle = parseBooleanEnv(explicitToggle);
  if (parsedToggle !== null) {
    return parsedToggle;
  }

  if (explicitToggle) {
    // Backward compatible: treat any non-empty custom value as enabled.
    return true;
  }

  // Auto-enable when OTP mailbox credentials are present.
  return Boolean(
    envValue(process.env.OTP_EMAIL_IMAP_HOST) &&
      envValue(process.env.OTP_EMAIL_USERNAME) &&
      envValue(process.env.OTP_EMAIL_PASSWORD)
  );
}

function getBetterLetterPrimaryCredentials() {
  const email =
    envValue(process.env.BETTERLETTER_USER_EMAIL) ||
    envValue(process.env.user_email) ||
    envValue(process.env.USER_EMAIL) ||
    envValue(process.env.ADMIN_PANEL_EMAIL);
  const password =
    envValue(process.env.BETTERLETTER_USER_PASSWORD) ||
    envValue(process.env.user_password) ||
    envValue(process.env.USER_PASSWORD) ||
    envValue(process.env.ADMIN_PANEL_PASSWORD);

  if (!email || !password) return null;
  return { email, password };
}

function isBetterLetterSignInUrl(url) {
  const lowerUrl = String(url || "").toLowerCase();
  return (
    lowerUrl.includes("/sign-in") ||
    lowerUrl.includes("/users/log_in") ||
    lowerUrl.includes("/login")
  );
}

async function inspectBetterLetterAuthSurface(page) {
  const url = page.url();
  const practicesVisible = await page
    .locator('a[href^="/admin_panel/practices/"]')
    .first()
    .isVisible({ timeout: 300 })
    .catch(() => false);

  const emailVisible = await page
    .locator(
      [
        'input[type="email"]',
        'input[name="email"]',
        'input[id*="email" i]',
        'input[autocomplete="username"]',
      ].join(", ")
    )
    .first()
    .isVisible({ timeout: 300 })
    .catch(() => false);

  const passwordVisible = await page
    .locator(
      [
        'input[type="password"]',
        'input[name="password"]',
        'input[id*="password" i]',
        'input[autocomplete="current-password"]',
      ].join(", ")
    )
    .first()
    .isVisible({ timeout: 300 })
    .catch(() => false);

  const codeVisible = await page
    .locator(
      [
        'input[autocomplete="one-time-code"]',
        'input[name*="code" i]',
        'input[id*="code" i]',
        'input[inputmode="numeric"]',
        'input[type="tel"]',
      ].join(", ")
    )
    .first()
    .isVisible({ timeout: 300 })
    .catch(() => false);

  const bodyText = await page.locator("body").innerText().catch(() => "");
  const bodyLower = String(bodyText || "").toLowerCase();

  const on2faByText =
    bodyLower.includes("security code") ||
    bodyLower.includes("one-time code") ||
    bodyLower.includes("verification code");

  const onLoginPage =
    isBetterLetterSignInUrl(url) ||
    (emailVisible && passwordVisible) ||
    bodyLower.includes("sign in to betterletter");

  const on2faPage = !practicesVisible && (codeVisible || (on2faByText && !passwordVisible));

  return {
    url,
    onLoginPage,
    on2faPage,
    practicesVisible,
  };
}

async function firstExistingLocator(locators = []) {
  for (const locator of locators) {
    const count = await locator.count().catch(() => 0);
    if (count > 0) return locator.first();
  }
  return null;
}

async function robustFillInput(locator, value) {
  const target = String(value || "");

  await locator.scrollIntoViewIfNeeded().catch(() => {});
  await locator.click({ timeout: 5000 }).catch(() => {});
  await locator.fill("").catch(() => {});

  await locator.type(target, { delay: 12 }).catch(async () => {
    await locator.fill(target).catch(() => {});
  });

  let actual = await locator.inputValue().catch(() => "");
  if (actual !== target) {
    await locator.fill(target).catch(() => {});
    actual = await locator.inputValue().catch(() => "");
  }

  return String(actual || "");
}

async function fillBetterLetterPrimaryLogin(page, email, password) {
  const emailField = await firstExistingLocator([
    page.getByLabel(/email/i),
    page
      .locator(
        [
          'input[type="email"]',
          'input[name="email"]',
          'input[id*="email" i]',
          'input[autocomplete="username"]',
        ].join(", ")
      )
      .first(),
  ]);

  const passwordField = await firstExistingLocator([
    page.getByLabel(/^password$/i),
    page
      .locator(
        [
          'input[type="password"]',
          'input[name="password"]',
          'input[id*="password" i]',
          'input[autocomplete="current-password"]',
        ].join(", ")
      )
      .first(),
  ]);

  if (!emailField || !passwordField) {
    return false;
  }

  const typedEmail = await robustFillInput(emailField, email);
  const typedPassword = await robustFillInput(passwordField, password);
  console.log(
    `✔ BetterLetter sign-in form filled (emailLen=${typedEmail.length}, passwordLen=${typedPassword.length})`
  );

  if (!typedEmail || !typedPassword) {
    throw new Error("Could not populate BetterLetter sign-in form fields reliably.");
  }

  const submit = page
    .locator(
      [
        'button:has-text("Sign in")',
        'button:has-text("Log in")',
        'button:has-text("SIGN IN")',
        'button[type="submit"]',
      ].join(", ")
    )
    .first();

  if ((await submit.count().catch(() => 0)) > 0) {
    await submit.click().catch(() => {});
  } else {
    await passwordField.press("Enter").catch(() => {});
  }

  return true;
}

async function waitForBetterLetter2faOrPractices(page, timeoutMs, options = {}) {
  const { allow2faReturn = true } = options;
  const startedAt = Date.now();
  let sawSignIn = false;
  let saw2fa = false;

  while (Date.now() - startedAt < timeoutMs) {
    const surface = await inspectBetterLetterAuthSurface(page);
    if (surface.practicesVisible || (!surface.onLoginPage && !surface.on2faPage)) return;
    if (surface.on2faPage) {
      saw2fa = true;
      if (allow2faReturn) return;
    }
    if (surface.onLoginPage) sawSignIn = true;
    await page.waitForTimeout(900);
  }

  await page.screenshot({ path: "betterletter-auth-timeout.png", fullPage: true }).catch(() => {});
  if (!allow2faReturn && saw2fa) {
    throw new Error(
      "BetterLetter 2FA did not complete in time after code submission. " +
        `Current URL: ${page.url()}. Screenshot: betterletter-auth-timeout.png`
    );
  }
  if (sawSignIn || isBetterLetterSignInUrl(page.url())) {
    throw new Error(
      "BetterLetter login did not advance from sign-in in time. " +
        `Current URL: ${page.url()}. Screenshot: betterletter-auth-timeout.png`
    );
  }

  throw new Error(
    "Timed out waiting for BetterLetter login to reach 2FA/practices. " +
      `Current URL: ${page.url()}. Screenshot: betterletter-auth-timeout.png`
  );
}

async function completeBetterLetter2fa(page, options = {}) {
  const auto2faFromEmail = Boolean(options.auto2faFromEmail);
  const authWaitTimeoutMs = Number(options.authWaitTimeoutMs) || 120000;
  const allowManualInput = options?.allowManualInput !== false;
  const maxAutoAttempts = Math.max(
    1,
    Number(process.env.BETTERLETTER_AUTO_2FA_MAX_ATTEMPTS || 2)
  );

  if (auto2faFromEmail) {
    try {
      let lastAutoError = null;

      for (let attempt = 1; attempt <= maxAutoAttempts; attempt += 1) {
        const requestedNewCode = await maybeRequestNewBetterLetter2faCode(page);
        const fetchStartedAt = new Date();
        console.log(
          requestedNewCode
            ? attempt === 1
              ? "🔑 Requested a new BetterLetter 2FA code. Fetching security code from email..."
              : `🔑 Requested another BetterLetter 2FA code (attempt ${attempt}/${maxAutoAttempts}). Fetching from email...`
            : attempt === 1
              ? "🔑 Fetching BetterLetter 2FA security code from email..."
              : `🔑 Retrying BetterLetter 2FA from email (attempt ${attempt}/${maxAutoAttempts})...`
        );

        const otpMessage = await fetchOtpCodeFromEmail(fetchStartedAt, {
          minReceivedAt: fetchStartedAt,
        });
        await fillAndSubmitBetterLetter2fa(page, otpMessage.code);
        await page.waitForLoadState("domcontentloaded", { timeout: 15000 }).catch(() => {});

        const result = await waitForBetterLetter2faCompletion(page, authWaitTimeoutMs);
        if (result.status === "success") {
          console.log("✔ BetterLetter 2FA code submitted automatically.");
          return;
        }

        lastAutoError = new Error(result.message);
        if (result.status === "invalid_code" && attempt < maxAutoAttempts) {
          console.log(`↻ BetterLetter rejected the OTP code. Retrying with a newer email code...`);
          continue;
        }

        throw lastAutoError;
      }

      if (lastAutoError) {
        throw lastAutoError;
      }
    } catch (error) {
      console.log(`⚠ Auto 2FA from email failed: ${error?.message || error}`);
      if (!allowManualInput) {
        throw createInteractiveAuthRequiredError(
          `Auto 2FA failed in headless mode: ${error?.message || String(error)}`
        );
      }
      if (!process.stdin.isTTY) {
        throw error;
      }
    }
  }

  if (!allowManualInput) {
    throw createInteractiveAuthRequiredError(
      "BetterLetter 2FA requires manual code entry, but browser is headless."
    );
  }

  console.log("\n🔐 BetterLetter is asking for a 2FA security code.");
  console.log("Enter the code in the browser window, then press ENTER here.\n");
  await waitForEnter();
}

async function maybeRequestNewBetterLetter2faCode(page) {
  const trigger = page
    .locator(
      [
        'a:has-text("generate new code")',
        'button:has-text("generate new code")',
        'a:has-text("new code")',
        'button:has-text("new code")',
      ].join(", ")
    )
    .first();

  if ((await trigger.count().catch(() => 0)) === 0) return false;

  try {
    await trigger.click({ timeout: 3000 });
    await page.waitForTimeout(1000);
    return true;
  } catch (_) {
    return false;
  }
}

function getBetterLetter2faInlineError(bodyText) {
  const text = String(bodyText || "").replace(/\s+/g, " ").trim();
  if (!text) return "";

  const matchers = [
    /the code is incorrect/i,
    /incorrect security code/i,
    /invalid security code/i,
    /invalid verification code/i,
    /verification code is invalid/i,
    /security code has expired/i,
    /code has expired/i,
  ];

  for (const matcher of matchers) {
    const match = text.match(matcher);
    if (match) return match[0];
  }

  return "";
}

function getOtpEmailConfig() {
  const host = envValue(process.env.OTP_EMAIL_IMAP_HOST);
  const username = envValue(process.env.OTP_EMAIL_USERNAME);
  const password = envValue(process.env.OTP_EMAIL_PASSWORD);

  if (!host || !username || !password) {
    throw new Error(
      "Missing OTP email IMAP config. Set OTP_EMAIL_IMAP_HOST, OTP_EMAIL_USERNAME, OTP_EMAIL_PASSWORD."
    );
  }

  const configuredMailboxes = uniqueList([
    ...parseEnvList(process.env.OTP_EMAIL_MAILBOX || "INBOX"),
    ...parseEnvList(process.env.OTP_EMAIL_MAILBOX_FALLBACKS),
    "INBOX",
    "BetterLetter2FA",
  ]);

  const subjectFilters = uniqueList(
    parseEnvList(process.env.OTP_EMAIL_SUBJECT_FILTER || "security code")
  );

  return {
    host,
    port: Number(process.env.OTP_EMAIL_IMAP_PORT || 993),
    secure: process.env.OTP_EMAIL_IMAP_SECURE
      ? envFlag(process.env.OTP_EMAIL_IMAP_SECURE)
      : true,
    username,
    password,
    mailboxes: configuredMailboxes.length ? configuredMailboxes : ["INBOX"],
    fromFilters: uniqueList(parseEnvList(process.env.OTP_EMAIL_FROM_FILTER)),
    subjectFilters: subjectFilters.length ? subjectFilters : ["security code"],
    lookbackMs: Number(process.env.OTP_EMAIL_LOOKBACK_MINUTES || 20) * 60 * 1000,
    timeoutMs: Number(process.env.OTP_EMAIL_TIMEOUT_SECONDS || 120) * 1000,
    pollMs: Number(process.env.OTP_EMAIL_POLL_SECONDS || 4) * 1000,
    codeRegex: process.env.OTP_EMAIL_CODE_REGEX || "\\b(\\d{6})\\b",
  };
}

function parseEnvList(value) {
  return String(value || "")
    .split(/[\n,;]+/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function uniqueList(values = []) {
  const out = [];
  const seen = new Set();
  for (const value of values) {
    const current = String(value || "").trim();
    if (!current) continue;
    const key = current.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(current);
  }
  return out;
}

function parseOtpFromRawMessage(raw, codeRegex) {
  const text = String(raw || "");
  const patterns = [];

  if (codeRegex) {
    const rawPattern = String(codeRegex);
    patterns.push(rawPattern);
    const unescaped = rawPattern.replace(/\\\\/g, "\\");
    if (unescaped !== rawPattern) patterns.push(unescaped);
  }

  patterns.push("\\b(\\d{6})\\b");

  for (const pattern of patterns) {
    try {
      const match = text.match(new RegExp(pattern, "i"));
      if (match) return match[1] || match[0] || "";
    } catch (_) {
      // Skip invalid custom regex and continue with fallback patterns.
    }
  }

  return "";
}

function resolveImapFlowClass() {
  const candidates = [];
  const explicitPath = envValue(process.env.OTP_EMAIL_IMAPFLOW_PATH);

  if (explicitPath) {
    candidates.push(explicitPath);
  }

  candidates.push("imapflow");

  for (const candidate of candidates) {
    try {
      const moduleRef = require(candidate);
      if (typeof moduleRef?.ImapFlow === "function") {
        return moduleRef.ImapFlow;
      }
    } catch (_) {}
  }

  throw new Error(
    "Could not load imapflow. Install it in this project (npm install imapflow) " +
      "or set OTP_EMAIL_IMAPFLOW_PATH to an existing imapflow module path."
  );
}

async function fetchOtpCodeFromEmail(sinceDate, options = {}) {
  const ImapFlow = resolveImapFlowClass();
  const cfg = getOtpEmailConfig();
  const client = new ImapFlow({
    host: cfg.host,
    port: cfg.port,
    secure: cfg.secure,
    auth: {
      user: cfg.username,
      pass: cfg.password,
    },
    logger: false,
  });

  const startedAt = Date.now();
  const loginMs = sinceDate instanceof Date ? sinceDate.getTime() : Date.now();
  const searchSince = new Date(Math.max(0, loginMs - cfg.lookbackMs));
  const minReceivedAt =
    options?.minReceivedAt instanceof Date ? options.minReceivedAt : sinceDate instanceof Date ? sinceDate : null;
  const minReceivedMs = minReceivedAt ? minReceivedAt.getTime() : 0;
  const minReceivedToleranceMs = Math.max(
    0,
    Number(process.env.OTP_EMAIL_MIN_RECEIVED_TOLERANCE_SECONDS || 15) * 1000
  );

  await client.connect();
  const candidateMailboxes = await resolveAvailableOtpMailboxes(client, cfg.mailboxes);
  console.log(
    `📨 OTP email polling: mailboxes=${candidateMailboxes.join("|") || "none"}, ` +
      `fromFilter=${cfg.fromFilters.join("|") || "none"}, ` +
      `subjectFilter=${cfg.subjectFilters.join("|") || "none"}, ` +
      `minReceivedAt=${minReceivedAt ? minReceivedAt.toISOString() : "none"}`
  );

  try {
    while (Date.now() - startedAt < cfg.timeoutMs) {
      for (const mailbox of candidateMailboxes) {
        const otpMessage = await findOtpInMailbox(client, {
          mailbox,
          searchSince,
          minReceivedMs,
          minReceivedToleranceMs,
          fromFilters: cfg.fromFilters,
          subjectFilters: cfg.subjectFilters,
          codeRegex: cfg.codeRegex,
        }).catch(() => null);

        if (otpMessage?.code) {
          return otpMessage;
        }
      }

      await client.noop().catch(() => {});
      await pageSleep(cfg.pollMs);
    }
  } finally {
    await client.logout().catch(() => {});
  }

  throw new Error(
    `Timed out waiting for OTP email after ${Math.floor(cfg.timeoutMs / 1000)} seconds ` +
      `(mailboxes=${candidateMailboxes.join("|") || "none"}, ` +
      `fromFilter=${cfg.fromFilters.join("|") || "none"}, ` +
      `subjectFilter=${cfg.subjectFilters.join("|") || "none"}, ` +
      `lookbackMinutes=${Math.floor(cfg.lookbackMs / 60000)}).`
  );
}

async function resolveAvailableOtpMailboxes(client, requestedMailboxes = []) {
  const requested = uniqueList(requestedMailboxes);
  if (!requested.length) return ["INBOX"];

  const discovered = [];
  try {
    for await (const box of client.list()) {
      const path = String(box?.path || box?.name || "").trim();
      if (path) discovered.push(path);
    }
  } catch (_) {
    return requested;
  }

  if (!discovered.length) return requested;

  const resolved = [];
  for (const mailbox of requested) {
    const lowerMailbox = mailbox.toLowerCase();
    let match = discovered.find((item) => item.toLowerCase() === lowerMailbox);

    if (!match) {
      match = discovered.find((item) => {
        const lowerItem = item.toLowerCase();
        return lowerItem.endsWith(`/${lowerMailbox}`) || lowerItem.includes(lowerMailbox);
      });
    }

    resolved.push(match || mailbox);
  }

  const likelyOtpMailboxes = discovered.filter((item) =>
    /(betterletter|2fa|otp|security)/i.test(item)
  );

  return uniqueList([...resolved, ...likelyOtpMailboxes]);
}

async function findOtpInMailbox(
  client,
  {
    mailbox,
    searchSince,
    minReceivedMs = 0,
    minReceivedToleranceMs = 0,
    fromFilters = [],
    subjectFilters = [],
    codeRegex = "",
  }
) {
  const lock = await client.getMailboxLock(mailbox);

  try {
    const uidSet = new Set();
    const rememberUids = (uids = []) => {
      for (const uid of uids) uidSet.add(uid);
    };

    if (fromFilters.length) {
      for (const fromFilter of fromFilters) {
        const query = { since: searchSince, from: fromFilter };
        const uids = await client.search(query).catch(() => []);
        rememberUids(uids);
      }
      if (!uidSet.size) {
        rememberUids(await client.search({ since: searchSince }).catch(() => []));
      }
    } else {
      rememberUids(await client.search({ since: searchSince }).catch(() => []));
    }

    const newestFirst = [...uidSet].sort((a, b) => b - a).slice(0, 40);

    for (const uid of newestFirst) {
      const msg = await client
        .fetchOne(uid, {
          envelope: true,
          internalDate: true,
          source: true,
        })
        .catch(() => null);

      if (!msg?.source) continue;
      if (msg.internalDate && msg.internalDate < searchSince) continue;
      if (
        minReceivedMs > 0 &&
        msg.internalDate &&
        msg.internalDate.getTime() + minReceivedToleranceMs < minReceivedMs
      ) {
        continue;
      }

      const subject = String(msg.envelope?.subject || "");
      if (!subjectMatchesAnyFilter(subject, subjectFilters)) continue;

      const raw = Buffer.from(msg.source).toString("utf8");
      const code = parseOtpFromRawMessage(raw, codeRegex);
      if (code) {
        return {
          code,
          mailbox,
          subject,
          internalDate: msg.internalDate || null,
        };
      }
    }
  } finally {
    lock.release();
  }

  return null;
}

function subjectMatchesAnyFilter(subject, subjectFilters = []) {
  if (!subjectFilters.length) return true;
  const lowerSubject = String(subject || "").toLowerCase();
  return subjectFilters.some((filter) =>
    lowerSubject.includes(String(filter || "").toLowerCase())
  );
}

async function fillAndSubmitBetterLetter2fa(page, code) {
  const codeInput = await firstExistingLocator([
    page.getByLabel(/security code/i),
    page.locator('input[autocomplete="one-time-code"]').first(),
    page
      .locator(
        [
          'input[name*="code" i]',
          'input[id*="code" i]',
          'input[inputmode="numeric"]',
          'input[type="tel"]',
        ].join(", ")
      )
      .first(),
  ]);

  if (!codeInput) {
    throw new Error("Could not find BetterLetter 2FA code input.");
  }

  const typedCode = await robustFillInput(codeInput, String(code));
  if (typedCode !== String(code)) {
    throw new Error("Could not populate BetterLetter 2FA code reliably.");
  }

  const verifyButton = page
    .locator(
      [
        'button:has-text("Verify Code")',
        'button:has-text("VERIFY CODE")',
        'button:has-text("Verify")',
        'button[type="submit"]',
      ].join(", ")
    )
    .first();

  if ((await verifyButton.count().catch(() => 0)) > 0) {
    await verifyButton.click().catch(async () => {
      await verifyButton.click({ force: true }).catch(() => {});
    });
  } else {
    await codeInput.press("Enter").catch(() => {});
  }
}

async function waitForBetterLetter2faCompletion(page, timeoutMs) {
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    const surface = await inspectBetterLetterAuthSurface(page);
    if (surface.practicesVisible || (!surface.onLoginPage && !surface.on2faPage)) {
      return { status: "success" };
    }

    const bodyText = await page.locator("body").innerText().catch(() => "");
    const inlineError = getBetterLetter2faInlineError(bodyText);
    if (inlineError) {
      return {
        status: "invalid_code",
        message:
          `BetterLetter rejected the submitted 2FA code (${inlineError}). ` +
          `Current URL: ${page.url()}`,
      };
    }

    await page.waitForTimeout(700);
  }

  await page.screenshot({ path: "betterletter-auth-timeout.png", fullPage: true }).catch(() => {});
  throw new Error(
    "BetterLetter 2FA did not complete in time after code submission. " +
      `Current URL: ${page.url()}. Screenshot: betterletter-auth-timeout.png`
  );
}

function pageSleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isCdpAttachConnectionError(error) {
  const text = String(error?.message || error || "").toLowerCase();
  return (
    text.includes("connectovercdp") ||
    text.includes("econnrefused") ||
    text.includes("retrieving websocket url") ||
    text.includes("failed to connect")
  );
}


/* ------------ Docman helpers ------------ */

async function ensureDocmanLoggedIn(page, { odsCode, adminUsername, adminPassword }, options = {}) {
  const navigationTimeoutMs = Number(options?.timeouts?.navigationMs) || 60000;
  const selectorTimeoutMs = Number(options?.timeouts?.selectorMs) || 60000;
  const fastCheckMs = Number(options?.timeouts?.docmanLoginFastCheckMs) || 3500;
  const deepCheckMs = Number(options?.timeouts?.docmanLoginDeepCheckMs) || 9000;
  const allowManualFallback = options?.allowManualFallback !== false;

  console.log("➡ Checking Docman session (attempting Filing directly)...");
  const resp = await page.goto(DOCMAN_FILING_URL, {
    waitUntil: "domcontentloaded",
    timeout: navigationTimeoutMs,
  });
  await page.waitForTimeout(150);

  const authState = await inspectDocmanAuthSurface(page, deepCheckMs);
  const onLoginPage = authState.onLoginPage;

  console.log(
    `[Docman auth check] status=${resp?.status?.() ?? "n/a"} url=${authState.url} onLoginPage=${onLoginPage}`
  );

  if (!onLoginPage) {
    console.log("✔ Docman appears already logged in (reused session).");
    return true;
  }

  console.log(`🔐 Docman login required. Logging in for ODS: ${odsCode}`);

  await tryOpenDocmanExplicitSignIn(page);

  const { orgField, userField, passField } = getDocmanLoginFieldLocators(page);

  // Wait for fields and fill (supports both OrganisationCode and OdsCode variants)
  await orgField.waitFor({ timeout: selectorTimeoutMs });
  await overwriteInput(orgField, odsCode, "Organisation Code");

  await userField.waitFor({ timeout: selectorTimeoutMs });
  await overwriteInput(userField, adminUsername, "User Name");

  await passField.waitFor({ timeout: selectorTimeoutMs });
  await overwriteInput(passField, adminPassword, "Password");

  // Prefer submit inside the password field's form to avoid clicking wrong buttons.
  const formScopedSubmit = passField
    .locator("xpath=ancestor::form[1]")
    .locator(
      [
        'button[type="submit"]',
        'input[type="submit"]',
        'button:has-text("Sign In")',
        'button:has-text("Sign in")',
      ].join(", ")
    )
    .first();

  const submitCount = await formScopedSubmit.count().catch(() => 0);
  if (submitCount > 0) {
    await formScopedSubmit.waitFor({ state: "attached", timeout: 30000 });
    await formScopedSubmit.click({ timeout: 30000 }).catch(async () => {
      await passField.press("Enter").catch(() => {});
    });
  } else {
    await passField.press("Enter").catch(() => {});
  }

  // Fast settle after submit: do not block on multiple long waits.
  await Promise.race([
    page.waitForURL(
      (url) => {
        const u = String(url).toLowerCase();
        return !u.includes("/account/login") && !u.includes("/account/prelogin");
      },
      { timeout: Math.min(10000, navigationTimeoutMs) }
    ).catch(() => null),
    page.waitForLoadState("domcontentloaded", { timeout: 6000 }).catch(() => null),
    page.waitForTimeout(1200),
  ]);

  // Fast-path verification first on current page.
  let postLoginState = await inspectDocmanAuthSurface(page, fastCheckMs);

  // If still on login/auth hand-off, force Filing and do a deeper check.
  if (postLoginState.onLoginPage) {
    await page.goto(DOCMAN_FILING_URL, { waitUntil: "commit", timeout: navigationTimeoutMs });
    postLoginState = await inspectDocmanAuthSurface(page, deepCheckMs);
  }
  const stillLogin = postLoginState.onLoginPage;

  if (stillLogin) {
    if (allowManualFallback && process.stdin.isTTY) {
      console.log("\n⚠ Docman automatic login did not complete.");
      console.log("Complete the Docman sign-in manually in the opened browser window.");
      console.log("Then press ENTER here to continue.\n");
      await waitForEnter();

      await page.goto(DOCMAN_FILING_URL, { waitUntil: "commit", timeout: navigationTimeoutMs }).catch(() => {});
      const manualPostState = await inspectDocmanAuthSurface(page, deepCheckMs);
      if (!manualPostState.onLoginPage) {
        console.log("✔ Docman login completed (manual fallback).");
        return true;
      }
    }

    const debugFiles = await captureDocmanLoginDebugFiles(page, "failed");
    throw new Error(
      "Docman login did not complete (still on login page after submitting). " +
      "This could be wrong creds, SSO restriction, or extra prompt. " +
      `Current URL: ${postLoginState.url}. ` +
      `Debug files: ${debugFiles.screenshotPath}, ${debugFiles.htmlPath}`
    );
  }

  console.log("✔ Docman login completed.");
  return true;
}

async function ensureDocmanSessionForPractice(page, {
  practiceName,
  odsCode,
  adminUsername,
  adminPassword,
}, options = {}) {
  const {
    forceFreshLogin = false,
    skipExistingSessionCheck = false,
    timeouts = {},
    allowManualFallback = true,
  } = options;

  if (forceFreshLogin) {
    console.log("↻ Forcing fresh Docman login (clearing any existing Docman session first).");
    await logoutDocman(page);
    await ensureDocmanLoggedIn(
      page,
      { odsCode, adminUsername, adminPassword },
      { timeouts, allowManualFallback }
    );

    const post = await getDocmanAuthState(page, { timeouts });
    if (!post.loggedIn) {
      throw new Error("Docman login failed after forcing a fresh session.");
    }
    if (post.orgName && !practiceMatches(practiceName, post.orgName)) {
      throw new Error(
        `Docman logged in, but organisation is "${post.orgName}" (expected "${practiceName}").`
      );
    }
    return;
  }

  if (skipExistingSessionCheck) {
    await ensureDocmanLoggedIn(
      page,
      { odsCode, adminUsername, adminPassword },
      { timeouts, allowManualFallback }
    );
    return;
  }

  const state = await getDocmanAuthState(page, { timeouts });

  if (!state.loggedIn) {
    await ensureDocmanLoggedIn(
      page,
      { odsCode, adminUsername, adminPassword },
      { timeouts, allowManualFallback }
    );
    return;
  }

  if (!state.orgName) {
    console.log(
      "⚠ Existing Docman session detected but organisation could not be confirmed after checks. Logging out immediately and performing a fresh login."
    );
    await logoutDocman(page);
    await ensureDocmanLoggedIn(
      page,
      { odsCode, adminUsername, adminPassword },
      { timeouts, allowManualFallback }
    );
    return;
  }

  if (practiceMatches(practiceName, state.orgName)) {
    console.log(`✔ Reusing Docman session for organisation: ${state.orgName}`);
    return;
  }

  console.log(
    `⚠ Docman session is for "${state.orgName}" but expected "${practiceName}". Resetting Docman login…`
  );

  await logoutDocman(page);
  await ensureDocmanLoggedIn(
    page,
    { odsCode, adminUsername, adminPassword },
    { timeouts, allowManualFallback }
  );

  const postLoginState = await getDocmanAuthState(page, { timeouts });
  if (postLoginState.loggedIn && postLoginState.orgName && !practiceMatches(practiceName, postLoginState.orgName)) {
    throw new Error(
      `Docman logged in, but organisation is still "${postLoginState.orgName}" (expected "${practiceName}").`
    );
  }
}

async function clearDocmanAuthFromContext(context) {
  const docmanUrls = [
    `${DOCMAN_ORIGIN}/`,
    DOCMAN_LOGIN_URL,
    DOCMAN_FILING_URL,
  ];

  let removedCount = 0;

  try {
    // Cookie-only reset to avoid affecting BetterLetter session state.
    const docmanCookies = await context.cookies(docmanUrls);
    if (!docmanCookies.length) return 0;

    const expiryCookies = docmanCookies
      .filter((cookie) => {
        const domain = String(cookie.domain || "").replace(/^\./, "").toLowerCase();
        return domain.endsWith(DOCMAN_HOST_SUFFIX);
      })
      .map((cookie) => {
        const expired = {
          name: cookie.name,
          value: "",
          domain: cookie.domain,
          path: cookie.path || "/",
          expires: 0,
          httpOnly: Boolean(cookie.httpOnly),
          secure: Boolean(cookie.secure),
        };

        if (
          cookie.sameSite === "Lax" ||
          cookie.sameSite === "Strict" ||
          cookie.sameSite === "None"
        ) {
          expired.sameSite = cookie.sameSite;
        }

        return expired;
      });

    removedCount = expiryCookies.length;
    if (expiryCookies.length) {
      await context.addCookies(expiryCookies).catch(() => {});
    }

    return removedCount;
  } catch (_) {
    return removedCount;
  }
}

async function tryOpenDocmanExplicitSignIn(page) {
  const trigger = page
    .locator(
      [
        'button:has-text("Sign in to Continue")',
        'a:has-text("Sign in to Continue")',
        'button:has-text("Sign In to Continue")',
        'a:has-text("Sign In to Continue")',
      ].join(", ")
    )
    .first();

  const visible = await trigger.isVisible({ timeout: 500 }).catch(() => false);
  if (!visible) return false;

  await trigger.click({ timeout: 5000 }).catch(() => {});
  await page.waitForTimeout(300);
  return true;
}

async function captureDocmanLoginDebugFiles(page, suffix = "failed") {
  const safe = String(suffix || "failed").replace(/[^a-z0-9_-]+/gi, "-");
  const screenshotPath = `docman-login-${safe}.png`;
  const htmlPath = `docman-login-${safe}.html`;

  await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});

  const html = await page.content().catch(() => "");
  if (html) {
    const fs = require("fs");
    fs.writeFileSync(htmlPath, html, "utf8");
  }

  return { screenshotPath, htmlPath };
}

async function overwriteInput(locator, value, label = "field") {
  await locator.click({ clickCount: 3 }).catch(() => {});
  await locator.press("ControlOrMeta+A").catch(() => {});
  await locator.press("Backspace").catch(() => {});
  await locator.fill("");
  await locator.type(value, { delay: 15 });

  const typed = await locator.inputValue().catch(() => "");
  if (typed !== value) {
    await locator.fill(value);
  }

  const finalValue = await locator.inputValue().catch(() => "");
  if (finalValue !== value) {
    throw new Error(`Docman ${label} did not stick in the input field.`);
  }
}

async function getDocmanAuthState(page, options = {}) {
  const navigationTimeoutMs = Number(options?.timeouts?.navigationMs) || 60000;
  const authCheckMs = Number(options?.timeouts?.docmanLoginDeepCheckMs) || 8000;

  await page.goto(DOCMAN_FILING_URL, { waitUntil: "domcontentloaded", timeout: navigationTimeoutMs });
  await page.waitForTimeout(300);
  const authSurface = await inspectDocmanAuthSurface(page, authCheckMs);
  const loggedIn = !authSurface.onLoginPage;
  if (!loggedIn) {
    return { loggedIn: false, orgName: null };
  }

  const orgName = await getCurrentDocmanOrgName(page);
  return { loggedIn: true, orgName };
}

async function getCurrentDocmanOrgName(page) {
  for (let attempt = 0; attempt < 3; attempt++) {
    const org = await detectDocmanOrgName(page);
    if (org) return org;

    // Some tenants only reveal org details after opening user/profile menu.
    await tryOpenDocmanUserMenu(page);
    const orgAfterMenu = await detectDocmanOrgName(page);
    if (orgAfterMenu) return orgAfterMenu;

    await page.waitForTimeout(250);
  }
  return null;
}

async function detectDocmanOrgName(page) {
  const selectors = [
    '[class*="user" i]:has-text("Docman System")',
    '[class*="profile" i]:has-text("Docman System")',
    '[class*="org" i]',
    '[id*="org" i]',
    "header :text-matches('Docman System', 'i')",
    'text=/Docman System/i',
  ];

  // Main page selectors.
  for (const selector of selectors) {
    const loc = page.locator(selector).first();
    const visible = await loc.isVisible({ timeout: 500 }).catch(() => false);
    if (!visible) continue;

    const text = (await loc.innerText().catch(() => "")) || "";
    const org = extractOrgFromText(text);
    if (org) return org;
  }

  // Fall back to frame text scanning.
  for (const frame of page.frames()) {
    const bodyText = await frame
      .locator("body")
      .first()
      .innerText()
      .catch(() => "");
    const org = extractOrgFromText(bodyText || "");
    if (org) return org;
  }

  return null;
}

function extractOrgFromText(text) {
  if (!text) return null;

  const compact = text.replace(/\s+/g, " ").trim();
  if (!compact) return null;

  const patterns = [
    /Docman System\s*-\s*([^|\n\r]+)/i,
    /Organisation\s*:\s*([^|\n\r]+)/i,
    /Organization\s*:\s*([^|\n\r]+)/i,
    /Practice\s*:\s*([^|\n\r]+)/i,
  ];

  for (const pattern of patterns) {
    const m = compact.match(pattern);
    if (!m) continue;
    const candidate = (m[1] || "").trim();
    if (candidate && candidate.length >= 2) return candidate;
  }

  return null;
}

async function tryOpenDocmanUserMenu(page) {
  const userMenuCandidates = [
    'button:has-text("User")',
    'a:has-text("User")',
    '[class*="user" i] button',
    '[class*="profile" i] button',
    '[class*="avatar" i]',
  ];

  for (const selector of userMenuCandidates) {
    const menu = page.locator(selector).first();
    const visible = await menu.isVisible({ timeout: 400 }).catch(() => false);
    if (!visible) continue;
    await menu.click().catch(() => {});
    await page.waitForTimeout(120);
    return true;
  }

  return false;
}

function practiceMatches(expectedPracticeName, docmanOrgName) {
  if (!expectedPracticeName || !docmanOrgName) return false;
  const normalize = (value) =>
    value
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9 ]+/g, " ")
      .replace(/\b(the|surgery|medical|practice|centre|center|health|clinic)\b/g, " ")
      .replace(/\s+/g, " ")
      .trim();

  const a = normalize(expectedPracticeName);
  const b = normalize(docmanOrgName);
  if (!a || !b) return false;
  if (a === b) return true;

  // allow strong partial match for short naming differences
  return a.includes(b) || b.includes(a);
}

async function logoutDocman(page) {
  // Best-effort sign out. Many Docman deployments support /Account/Logout.
  // Even if it doesn't, we still handle the result safely.
  await page.goto(DOCMAN_LOGOUT_URL, {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  }).catch(() => {});

  // If logout route isn't supported, try clicking User menu → Logout (best-effort)
  const userMenu = page.locator('text=/User/i').first();
  const userVisible = await userMenu.isVisible({ timeout: 1500 }).catch(() => false);
  if (userVisible) {
    await userMenu.click().catch(() => {});
    await page.locator('text=/Log out|Logout|Sign out/i').first().click().catch(() => {});
  }

  // After logout, we should end up on login page
  await page.waitForTimeout(500);
}


async function gotoDocmanFilingAndActivate(page, options = {}) {
  const { skipDialogCheck = false, timeouts = {} } = options;
  const navigationTimeoutMs = Number(timeouts?.navigationMs) || 60000;
  const authCheckMs = Number(timeouts?.docmanLoginDeepCheckMs) || 8000;

  console.log("➡ Navigating to Filing");
  await page.goto(DOCMAN_FILING_URL, {
    waitUntil: "domcontentloaded",
    timeout: navigationTimeoutMs,
  });

  const filingAuthState = await inspectDocmanAuthSurface(page, authCheckMs);
  if (filingAuthState.onLoginPage) {
    await page.screenshot({ path: "docman-filing-login-page.png", fullPage: true }).catch(() => {});
    throw new Error(
      "Docman is still on the login page while trying to open Filing. " +
      `Current URL: ${filingAuthState.url}. Screenshot: docman-filing-login-page.png`
    );
  }

  // Give SPAs time to render
  await page.waitForLoadState("networkidle").catch(() => {});

  // Close "Restore pages?" popup if it blocks clicks (best-effort)
  await dismissRestorePagesPopup(page);

  if (!skipDialogCheck) {
    await waitAndDismissBlockingDialogs(page, "after filing navigation");
  }

  // Filing UI might be in an iframe — pick the best frame
  const frame = await getDocmanFilingFrame(page);

  // Activate Filing screen (best-effort)
  const allDocsBtn = frame.locator("span.all-docs-count").first();
  if (await allDocsBtn.count()) {
    await allDocsBtn.click().catch(() => {});
  }

  // ✅ Wait for folders pane using multiple possible selectors
  const folderCandidates = [
    "#folders_list",
    "#folders",
    '[id*="folder" i]',
    '[class*="folder" i]',
    'text=/Folders/i',
  ];

  try {
    await waitForAnySelector(frame, folderCandidates, navigationTimeoutMs);
  } catch (e) {
    // 🔎 Debug dump to see what page we’re really on
    const url = page.url();
    const title = await page.title().catch(() => "");
    const bodyText = (await frame.locator("body").innerText().catch(() => "")) || "";

    await page.screenshot({ path: "docman-filing-timeout.png", fullPage: true }).catch(() => {});

    console.error("\n--- DOCMAN FILING DEBUG ---");
    console.error("URL:", url);
    console.error("TITLE:", title);
    console.error("BODY (first 400 chars):", bodyText.slice(0, 400));
    console.error("Saved screenshot: docman-filing-timeout.png");
    console.error("---------------------------\n");

    throw new Error(
      "Docman Filing did not render the folders pane (selectors not found). " +
      "See docman-filing-timeout.png + debug output above."
    );
  }

  console.log("✔ Docman Filing ready");
}

async function dismissRestorePagesPopup(page) {
  const popup = page.locator('text=/Restore pages\\?/i').first();
  const visible = await popup.isVisible({ timeout: 800 }).catch(() => false);
  if (!visible) return;

  // Try close button (X)
  await page.keyboard.press("Escape").catch(() => {});
  await page.locator('button[aria-label="Close"], button:has-text("×")').first().click().catch(() => {});
}

async function getDocmanFilingFrame(page) {
  // If Docman uses iframes, the Filing UI will be inside one.
  const frames = page.frames();

  // Prefer a frame that looks like DocumentViewer content
  const candidate =
    frames.find((f) => /DocumentViewer/i.test(f.url())) ||
    frames.find((f) => /docman/i.test(f.url())) ||
    page.mainFrame();

  return candidate;
}

function getDocmanLoginFieldLocators(page) {
  return {
    orgField: page
      .locator(
        [
          "#OrganisationCode",
          "#OrganizationCode",
          "#OdsCode",
          'input[name="OrganisationCode"]',
          'input[name="OrganizationCode"]',
          'input[name="OdsCode"]',
          'input[name*="organisation" i]',
          'input[name*="organization" i]',
          'input[placeholder*="organisation" i]',
          'input[placeholder*="organization" i]',
        ].join(", ")
      )
      .first(),
    userField: page
      .locator(
        [
          "#UserName",
          "#Username",
          'input[name="UserName"]',
          'input[name="Username"]',
          'input[name*="user" i]',
          'input[autocomplete="username"]',
        ].join(", ")
      )
      .first(),
    passField: page
      .locator(
        [
          "#Password",
          'input[name="Password"]',
          'input[type="password"]',
          'input[autocomplete="current-password"]',
        ].join(", ")
      )
      .first(),
  };
}

async function inspectDocmanAuthSurface(page, timeoutMs = 6000) {
  const start = Date.now();

  while (Date.now() - start < timeoutMs) {
    const url = page.url();
    const lowerUrl = url.toLowerCase();

    const { orgField, userField, passField } = getDocmanLoginFieldLocators(page);

    const loginByUrl =
      lowerUrl.includes("/account/login") ||
      lowerUrl.includes("/account/prelogin");

    const signInHeadingVisible = await page
      .locator("text=/Sign in to Continue/i")
      .first()
      .isVisible({ timeout: 250 })
      .catch(() => false);

    const autoSignInFailedVisible = await page
      .locator("text=/automatic sign-in failed/i")
      .first()
      .isVisible({ timeout: 250 })
      .catch(() => false);

    const orgVisible = await orgField.isVisible({ timeout: 250 }).catch(() => false);
    const userVisible = await userField.isVisible({ timeout: 250 }).catch(() => false);
    const passVisible = await passField.isVisible({ timeout: 250 }).catch(() => false);

    const filingUiVisible = await page
      .locator(
        [
          "span.all-docs-count",
          "#folders_list",
          "#folders",
          '[id*="folder" i]',
          '[class*="folder" i]',
          "text=/Filing/i",
        ].join(", ")
      )
      .first()
      .isVisible({ timeout: 250 })
      .catch(() => false);

    const onLoginPage =
      loginByUrl ||
      signInHeadingVisible ||
      autoSignInFailedVisible ||
      (orgVisible && (userVisible || passVisible));

    if (onLoginPage || filingUiVisible) {
      return { onLoginPage, url };
    }

    await page.waitForTimeout(250);
  }

  const timeoutUrl = page.url();
  const timeoutLower = timeoutUrl.toLowerCase();
  return {
    onLoginPage:
      timeoutLower.includes("/account/login") || timeoutLower.includes("/account/prelogin"),
    url: timeoutUrl,
  };
}

async function waitForAnySelector(frame, selectors, timeoutMs) {
  const start = Date.now();
  let lastErr = null;

  while (Date.now() - start < timeoutMs) {
    for (const sel of selectors) {
      try {
        const loc = frame.locator(sel).first();
        if (await loc.count()) {
          // If it exists, also ensure it’s attached/visible-ish
          await loc.waitFor({ state: "attached", timeout: 1500 }).catch(() => {});
          return sel;
        }
      } catch (e) {
        lastErr = e;
      }
    }
    await frame.page().waitForTimeout(300);
  }

  throw lastErr || new Error("No selectors matched in time");
}


/* ------------ helpers ------------ */

function printPhaseBanner(title, lines = []) {
  const width = 62;
  const bar = "=".repeat(width);
  console.log(`\n${bar}`);
  console.log(` ${title}`);
  for (const line of lines) {
    console.log(` - ${line}`);
  }
  console.log(`${bar}\n`);
}

function waitForEnter() {
  return new Promise((resolve) => {
    process.stdin.resume();
    process.stdin.once("data", () => resolve());
  });
}

async function waitAndDismissBlockingDialogs(
  page,
  reason = "unknown",
  windowMs = 5000,
  pollMs = 250
) {
  console.log(`⏳ Watching for blocking dialogs (${reason})`);

  const start = Date.now();
  let dismissedAny = false;
  let lastModalSeenAt = null;
  const quietExitMs = 1200;

  while (Date.now() - start < windowMs) {
    const modal = page.locator(".alertify.ajs-in, .alertify.ajs-fade.ajs-in");

    if (await modal.count()) {
      dismissedAny = true;
      lastModalSeenAt = Date.now();
      console.log("⚠ Blocking dialog detected — dismissing");

      const btn = modal
        .locator("button, a")
        .filter({ hasText: /ok|confirm|close|continue|yes|got it|×/i })
        .first();

      if (await btn.count()) {
        await btn.click({ force: true }).catch(() => {});
      } else {
        await page.keyboard.press("Escape").catch(() => {});
      }
    } else {
      const now = Date.now();
      if (!dismissedAny && now - start >= quietExitMs) {
        break;
      }
      if (dismissedAny && lastModalSeenAt && now - lastModalSeenAt >= quietExitMs) {
        break;
      }
    }

    await page.waitForTimeout(pollMs);
  }

  if (dismissedAny) console.log("✔ Dialog check complete");
  else console.log("✔ No blocking dialogs appeared");
}

module.exports = bootstrapDocmanSession;
module.exports.gotoDocmanFilingAndActivate = gotoDocmanFilingAndActivate;
