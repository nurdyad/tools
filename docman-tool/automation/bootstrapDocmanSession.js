// automation/bootstrapDocmanSession.js
const { getBrowserSession } = require("./browserSession");
const { fetchDocmanCreds } = require("./fetchDocmanCreds");

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
  } = sessionOptions;
  const { context, page } = await getBrowserSession(sessionOptions);

  // Always start on BetterLetter admin panel when the browser opens.
  await page.goto(BETTERLETTER_PRACTICES_URL, {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  }).catch(() => {});

  const planLines = [
    "1) Session health check",
    "2) Login to BetterLetter (manual username/password/2FA if prompted)",
    "3) Read Docman credentials from BetterLetter",
    "4) Login to Docman with BetterLetter credentials",
  ];
  if (!skipPostLoginDialogWatch) {
    planLines.push("5) Dismiss blocking dialogs");
  }

  printPhaseBanner(
    "DOCMAN BOOTSTRAP",
    planLines
  );

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

  // 1) Session health check (prints current auth status)
  printPhaseBanner("Step 1", ["Session health check"]);
  await sessionHealthCheck(page, { includeDocmanCheck: includeDocmanInHealthCheck });

  // 2) BetterLetter (persistent session)
  printPhaseBanner("Step 2", ["Login to BetterLetter"]);
  await ensureBetterLetterLoggedIn(page);

  // 3) Fetch Docman creds from BetterLetter
  printPhaseBanner("Step 3", ["Read Docman ODS/username/password from BetterLetter"]);
  const { odsCode, adminUsername, adminPassword } = await fetchDocmanCreds(
    page,
    practiceName.trim()
  );

  // 4) Ensure Docman is logged in and scoped to the target practice where possible.
  printPhaseBanner("Step 4", ["Login to Docman with BetterLetter credentials"]);
  await ensureDocmanSessionForPractice(page, {
    practiceName: practiceName.trim(),
    odsCode,
    adminUsername,
    adminPassword,
  }, {
    forceFreshLogin: forceFreshDocmanLogin,
    skipExistingSessionCheck: resetDocmanAuthAtStart,
  });

  // 5) Clear any post-login modal(s) when requested by workflow.
  if (!skipPostLoginDialogWatch) {
    printPhaseBanner("Step 5", ["Dismiss blocking dialogs"]);
    await waitAndDismissBlockingDialogs(page, "after docman login");
  }

  return {
    context,
    page,
    odsCode,
    adminUsername,
    adminPassword,
  };
}

/* ------------ session health check ------------ */

async function sessionHealthCheck(page, options = {}) {
  console.log("\n================ SESSION HEALTH CHECK ================");
  const includeDocmanCheck = Boolean(options.includeDocmanCheck);

  // ---------------- BetterLetter check ----------------
  let blLoggedIn = false;
  try {
    const resp = await page.goto("https://app.betterletter.ai/admin_panel/practices", {
      waitUntil: "domcontentloaded",
      timeout: 45000,
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
        timeout: 45000,
      });

      // Give redirects a moment
      await page.waitForTimeout(300);

      const authSurface = await inspectDocmanAuthSurface(page, 6000);
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
      timeout: 45000,
    });
  } catch (_) {}
}


/* ------------ BetterLetter helpers ------------ */

async function ensureBetterLetterLoggedIn(page) {
  await page.goto(BETTERLETTER_PRACTICES_URL, {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });

  // New BetterLetter login route
  const isLoginPage =
    page.url().includes("/sign-in") ||
    page.url().includes("/users/log_in") ||
    page.url().includes("/login") ||
    (await page.locator('input[type="email"], input[name="email"]').first()
      .isVisible({ timeout: 800 })
      .catch(() => false));

  // Unauthorized page body check (basic auth failure / access denied)
  const bodyText = ((await page.textContent("body").catch(() => "")) || "").trim();
  if (/unauthorized/i.test(bodyText)) {
    throw new Error(
      "BetterLetter returned Unauthorized (HTTP Basic Auth failed). Check basic auth creds in run.js."
    );
  }

  if (isLoginPage) {
    console.log("\n🔐 Please log into BetterLetter in the opened browser window.");
    console.log("When finished, come back here and press ENTER.\n");
    await waitForEnter();

    // After manual login, re-open practices
    await page.goto(BETTERLETTER_PRACTICES_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60000,
    });

    // Wait until we can see at least one practice link (Phoenix patch link)
    await page.waitForSelector('a[href^="/admin_panel/practices/"]', {
      timeout: 60000,
    });

    // If still on sign-in, fail clearly
    if (page.url().includes("/sign-in")) {
      throw new Error("BetterLetter still shows sign-in after manual login.");
    }
  } else {
    // Not a login page — still ensure practices list is present
    await page.waitForSelector('a[href^="/admin_panel/practices/"]', {
      timeout: 60000,
    });
  }

  return true;
}


/* ------------ Docman helpers ------------ */

async function ensureDocmanLoggedIn(page, { odsCode, adminUsername, adminPassword }) {
  console.log("➡ Checking Docman session (attempting Filing directly)...");
  const resp = await page.goto(DOCMAN_FILING_URL, {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });
  await page.waitForTimeout(150);

  const authState = await inspectDocmanAuthSurface(page, 8000);
  const onLoginPage = authState.onLoginPage;

  console.log(
    `[Docman auth check] status=${resp?.status?.() ?? "n/a"} url=${authState.url} onLoginPage=${onLoginPage}`
  );

  if (!onLoginPage) {
    console.log("✔ Docman appears already logged in (reused session).");
    return true;
  }

  console.log(`🔐 Docman login required. Logging in for ODS: ${odsCode}`);

  const { orgField, userField, passField } = getDocmanLoginFieldLocators(page);

  // Wait for fields and fill (supports both OrganisationCode and OdsCode variants)
  await orgField.waitFor({ timeout: 60000 });
  await overwriteInput(orgField, odsCode, "Organisation Code");

  await userField.waitFor({ timeout: 60000 });
  await overwriteInput(userField, adminUsername, "User Name");

  await passField.waitFor({ timeout: 60000 });
  await overwriteInput(passField, adminPassword, "Password");

  // Click submit
  const submit = page.locator('button[type="submit"], button:has-text("Sign In")').first();
  await submit.waitFor({ timeout: 30000 });
  await submit.click({ timeout: 30000 });

  // Fast settle after submit: do not block on multiple long waits.
  await Promise.race([
    page.waitForURL(
      (url) => {
        const u = String(url).toLowerCase();
        return !u.includes("/account/login") && !u.includes("/account/prelogin");
      },
      { timeout: 10000 }
    ).catch(() => null),
    page.waitForLoadState("domcontentloaded", { timeout: 6000 }).catch(() => null),
    page.waitForTimeout(1200),
  ]);

  // Fast-path verification first on current page.
  let postLoginState = await inspectDocmanAuthSurface(page, 3500);

  // If still on login/auth hand-off, force Filing and do a deeper check.
  if (postLoginState.onLoginPage) {
    await page.goto(DOCMAN_FILING_URL, { waitUntil: "commit", timeout: 60000 });
    postLoginState = await inspectDocmanAuthSurface(page, 9000);
  }
  const stillLogin = postLoginState.onLoginPage;

  if (stillLogin) {
    await page.screenshot({ path: "docman-login-failed.png", fullPage: true }).catch(() => {});
    throw new Error(
      "Docman login did not complete (still on login page after submitting). " +
      "This could be wrong creds, SSO restriction, or extra prompt. " +
      `Current URL: ${postLoginState.url}. Screenshot: docman-login-failed.png`
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
  } = options;

  if (forceFreshLogin) {
    console.log("↻ Forcing fresh Docman login (clearing any existing Docman session first).");
    await logoutDocman(page);
    await ensureDocmanLoggedIn(page, { odsCode, adminUsername, adminPassword });

    const post = await getDocmanAuthState(page);
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
    await ensureDocmanLoggedIn(page, { odsCode, adminUsername, adminPassword });
    return;
  }

  const state = await getDocmanAuthState(page);

  if (!state.loggedIn) {
    await ensureDocmanLoggedIn(page, { odsCode, adminUsername, adminPassword });
    return;
  }

  if (!state.orgName) {
    console.log(
      "⚠ Existing Docman session detected but organisation could not be confirmed after checks. Logging out immediately and performing a fresh login."
    );
    await logoutDocman(page);
    await ensureDocmanLoggedIn(page, { odsCode, adminUsername, adminPassword });
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
  await ensureDocmanLoggedIn(page, { odsCode, adminUsername, adminPassword });

  const postLoginState = await getDocmanAuthState(page);
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
    // Server-side session reset without changing the visible page.
    await context.request.get(DOCMAN_LOGOUT_URL, {
      timeout: 15000,
      failOnStatusCode: false,
    }).catch(() => {});

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

    // Final best-effort logout after cookie expiry.
    await context.request.get(DOCMAN_LOGOUT_URL, {
      timeout: 15000,
      failOnStatusCode: false,
    }).catch(() => {});

    return removedCount;
  } catch (_) {
    return removedCount;
  }
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

async function getDocmanAuthState(page) {
  await page.goto(DOCMAN_FILING_URL, { waitUntil: "domcontentloaded", timeout: 60000 });
  await page.waitForTimeout(300);
  const authSurface = await inspectDocmanAuthSurface(page, 8000);
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
  const { skipDialogCheck = false } = options;

  console.log("➡ Navigating to Filing");
  await page.goto(DOCMAN_FILING_URL, {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });

  const filingAuthState = await inspectDocmanAuthSurface(page, 8000);
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
    await waitForAnySelector(frame, folderCandidates, 60000);
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
