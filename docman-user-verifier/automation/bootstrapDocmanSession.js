// automation/bootstrapDocmanSession.js
const { getBrowserSession } = require("./browserSession");
const { fetchDocmanCreds } = require("./fetchDocmanCreds");

async function bootstrapDocmanSession(practiceName, sessionOptions = {}) {
  const { context, page } = await getBrowserSession(sessionOptions);

  // Session health check (prints current auth status)
  await sessionHealthCheck(page);

  // 1) BetterLetter (persistent session)
  await ensureBetterLetterLoggedIn(page);

  // 2) Fetch Docman creds from BetterLetter
  const { odsCode, adminUsername, adminPassword } = await fetchDocmanCreds(
    page,
    practiceName
  );

  // 3) Ensure Docman is logged in (reuse if still valid)
  await ensureDocmanLoggedIn(page, { odsCode, adminUsername, adminPassword });

  // ensure Docman is logged into the correct org/practice
  const currentOrg = await getCurrentDocmanOrgName(page);
  if (currentOrg && !practiceMatches(practiceName, currentOrg)) {
    console.log(`⚠ Docman currently in org "${currentOrg}" but expected "${practiceName}". Re-authing…`);
    await logoutDocman(page);
    await ensureDocmanLoggedIn(page, { odsCode, adminUsername, adminPassword });
  }

  // 4) Clear any post-login modal(s)
  await waitAndDismissBlockingDialogs(page, "after docman login");

  // 5) Navigate to Filing (and activate)
  await gotoDocmanFilingAndActivate(page);

  return {
    context,
    page,
    odsCode,
    adminUsername,
    adminPassword,
  };
}

/* ------------ session health check ------------ */

async function sessionHealthCheck(page) {
  console.log("\n================ SESSION HEALTH CHECK ================");

  const returnUrl = page.url();

  // ---------------- BetterLetter check ----------------
  let blLoggedIn = false;
  try {
    const resp = await page.goto("https://app.betterletter.ai/admin_panel/practices", {
      waitUntil: "domcontentloaded",
      timeout: 45000,
    });

    const status = resp?.status?.();
    const bodyText = (await page.textContent("body").catch(() => "")) || "";
    const unauthorized =
      status === 401 ||
      status === 403 ||
      /unauthorized/i.test(bodyText);

    const u = page.url();
    const redirectedToLogin = u.includes("/users/log_in") || u.includes("/login");

    blLoggedIn = !unauthorized && !redirectedToLogin;

    console.log(
      `BetterLetter debug: status=${status ?? "n/a"} unauthorized=${unauthorized} url=${u}`
    );
  } catch (e) {
    console.log("BetterLetter: ⚠ Unable to determine (navigation issue)");
  }

  console.log(`BetterLetter: ${blLoggedIn ? "✅ Logged in" : "❌ Not logged in"}`);

  // ---------------- Docman check ----------------
  let dmLoggedIn = false;
  try {
    const filingUrl = "https://production.docman.thirdparty.nhs.uk/DocumentViewer/Filing";
    const resp = await page.goto(filingUrl, {
      waitUntil: "domcontentloaded",
      timeout: 45000,
    });

    // Give redirects a moment
    await page.waitForTimeout(300);

    const status = resp?.status?.();
    const url = page.url();

    const loginTextVisible = await page
      .locator("text=/Sign in to Continue/i")
      .first()
      .isVisible({ timeout: 1200 })
      .catch(() => false);

    const orgFieldVisible = await page
      .locator('#OrganisationCode, #OrganizationCode, #OdsCode, input[name="OrganisationCode"], input[name="OrganizationCode"], input[name="OdsCode"]')
      .first()
      .isVisible({ timeout: 1200 })
      .catch(() => false);

    const userFieldVisible = await page
      .locator('#UserName, #Username, input[name="UserName"], input[name="Username"]')
      .first()
      .isVisible({ timeout: 1200 })
      .catch(() => false);

    const passFieldVisible = await page
      .locator('#Password, input[name="Password"], input[type="password"]')
      .first()
      .isVisible({ timeout: 1200 })
      .catch(() => false);

    const onLoginPage =
      url.includes("/Account/Login") ||
      loginTextVisible ||
      (orgFieldVisible && userFieldVisible && passFieldVisible);

    dmLoggedIn = !onLoginPage;

    console.log(
      `Docman debug: status=${status ?? "n/a"} url=${url} onLoginPage=${onLoginPage}`
    );
  } catch (e) {
    console.log("Docman: ⚠ Unable to determine (navigation issue)");
  }

  console.log(`Docman: ${dmLoggedIn ? "✅ Logged in" : "❌ Not logged in"}`);

  console.log("======================================================\n");

  // Restore the prior page (best effort)
  try {
    if (returnUrl && returnUrl !== "about:blank") {
      await page.goto(returnUrl, {
        waitUntil: "domcontentloaded",
        timeout: 45000,
      });
    }
  } catch (_) {}
}


/* ------------ BetterLetter helpers ------------ */

async function ensureBetterLetterLoggedIn(page) {
  await page.goto("https://app.betterletter.ai/admin_panel/practices", {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });

  if (page.url().includes("/users/log_in") || page.url().includes("/login")) {
    console.log("\n🔐 Please log into BetterLetter in the opened browser window.");
    console.log("This is a one-time login (saved in the persistent profile).");
    console.log("Then come back here and press ENTER.\n");
    await waitForEnter();

    await page.goto("https://app.betterletter.ai/admin_panel/practices", {
      waitUntil: "domcontentloaded",
      timeout: 60000,
    });

    if (page.url().includes("/users/log_in") || page.url().includes("/login")) {
      throw new Error(
        "BetterLetter still shows login screen after manual login. Login may not have completed."
      );
    }
  }

  return true;
}

/* ------------ Docman helpers ------------ */

async function ensureDocmanLoggedIn(page, { odsCode, adminUsername, adminPassword }) {
  const filingUrl =
    "https://production.docman.thirdparty.nhs.uk/DocumentViewer/Filing";

  console.log("➡ Checking Docman session (attempting Filing directly)...");
  const resp = await page.goto(filingUrl, { waitUntil: "domcontentloaded", timeout: 60000 });

  // Give SPA/login redirects a moment to settle
  await page.waitForTimeout(500);

  const url = page.url();
  const status = resp?.status?.();

  // ✅ Detect login by URL OR by visible login form text/fields
  const loginTextVisible = await page
    .locator("text=/Sign in to Continue/i")
    .first()
    .isVisible({ timeout: 1500 })
    .catch(() => false);

  const orgField = page.locator(
    '#OrganisationCode, #OrganizationCode, #OdsCode, input[name="OrganisationCode"], input[name="OrganizationCode"], input[name="OdsCode"]'
  ).first();

  const userField = page.locator(
    '#UserName, #Username, input[name="UserName"], input[name="Username"]'
  ).first();

  const passField = page.locator(
    '#Password, input[name="Password"], input[type="password"]'
  ).first();

  const orgVisible = await orgField.isVisible({ timeout: 1500 }).catch(() => false);
  const userVisible = await userField.isVisible({ timeout: 1500 }).catch(() => false);
  const passVisible = await passField.isVisible({ timeout: 1500 }).catch(() => false);

  const onLoginPage =
    url.includes("/Account/Login") ||
    loginTextVisible ||
    (orgVisible && userVisible && passVisible);

  console.log(
    `[Docman auth check] status=${status ?? "n/a"} url=${url} onLoginPage=${onLoginPage}`
  );

  if (!onLoginPage) {
    console.log("✔ Docman appears already logged in (reused session).");
    return true;
  }

  console.log(`🔐 Docman login required. Logging in for ODS: ${odsCode}`);

  // Wait for fields and fill (supports both OrganisationCode and OdsCode variants)
  await orgField.waitFor({ timeout: 60000 });
  await orgField.fill(odsCode);

  await userField.waitFor({ timeout: 60000 });
  await userField.fill(adminUsername);

  await passField.waitFor({ timeout: 60000 });
  await passField.fill(adminPassword);

  // Click submit
  const submit = page.locator('button[type="submit"], button:has-text("Sign In")').first();
  await Promise.all([
    submit.click(),
    page.waitForLoadState("domcontentloaded"),
  ]);

  // After submit, Docman may land somewhere else; ensure Filing is loaded
  await page.goto(filingUrl, { waitUntil: "domcontentloaded", timeout: 60000 });

  // Final verification: if still login, fail clearly
  const stillLogin =
    page.url().includes("/Account/Login") ||
    (await page.locator("text=/Sign in to Continue/i").first().isVisible({ timeout: 1500 }).catch(() => false));

  if (stillLogin) {
    throw new Error(
      "Docman login did not complete (still on login page after submitting). " +
      "This could be wrong creds, SSO restriction, or extra prompt."
    );
  }

  console.log("✔ Docman login completed.");
  return true;
}

async function getCurrentDocmanOrgName(page) {
  // In your screenshot it looks like:
  // "Mr Dyad Betterletter (Docman System Administrator) - HEATHVIEW MEDICAL PRACTICE"
  // We'll grab text containing "Docman System" then take the part after " - ".
  const header = page.locator('text=/Docman System/i').first();
  const visible = await header.isVisible({ timeout: 1500 }).catch(() => false);
  if (visible) {
    const t = (await header.innerText().catch(() => "")) || "";
    const idx = t.lastIndexOf(" - ");
    if (idx !== -1) return t.slice(idx + 3).trim();
  }

  // Fallback: scan body text for " - SOME PRACTICE"
  const body = (await page.textContent("body").catch(() => "")) || "";
  const m = body.match(/-\s*([A-Z0-9][A-Z0-9 \-']{3,})/);
  return m ? m[1].trim() : null;
}

function practiceMatches(expectedPracticeName, docmanOrgName) {
  if (!expectedPracticeName || !docmanOrgName) return false;
  const a = expectedPracticeName.trim().toLowerCase();
  const b = docmanOrgName.trim().toLowerCase();

  // allow partial match either direction (Alrewas vs ALREWAS SURGERY)
  return a.includes(b) || b.includes(a);
}

async function logoutDocman(page) {
  // Best-effort sign out. Many Docman deployments support /Account/Logout.
  // Even if it doesn't, we still handle the result safely.
  await page.goto("https://production.docman.thirdparty.nhs.uk/Account/Logout", {
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


async function gotoDocmanFilingAndActivate(page) {
  const docmanOrigin = "https://production.docman.thirdparty.nhs.uk";

  console.log("➡ Navigating to Filing");
  await page.goto(`${docmanOrigin}/DocumentViewer/Filing`, {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });

  // Give SPAs time to render
  await page.waitForLoadState("networkidle").catch(() => {});

  // Close "Restore pages?" popup if it blocks clicks (best-effort)
  await dismissRestorePagesPopup(page);

  await waitAndDismissBlockingDialogs(page, "after filing navigation");

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

function waitForEnter() {
  return new Promise((resolve) => {
    process.stdin.resume();
    process.stdin.once("data", () => resolve());
  });
}

async function waitAndDismissBlockingDialogs(
  page,
  reason = "unknown",
  windowMs = 10000,
  pollMs = 500
) {
  console.log(`⏳ Watching for blocking dialogs (${reason})`);

  const start = Date.now();
  let dismissedAny = false;

  while (Date.now() - start < windowMs) {
    const modal = page.locator(".alertify.ajs-in, .alertify.ajs-fade.ajs-in");

    if (await modal.count()) {
      dismissedAny = true;
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
    }

    await page.waitForTimeout(pollMs);
  }

  if (dismissedAny) console.log("✔ Dialog check complete");
  else console.log("✔ No blocking dialogs appeared");
}

module.exports = bootstrapDocmanSession;
