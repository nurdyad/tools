// verifyDocmanUsers.js
async function verifyDocmanUsers({ page, usernames }) {
  if (!page) {
    throw new Error("verifyDocmanUsers requires an existing Playwright page.");
  }

  console.log("Checking Docman users in existing session.");

  const authBefore = await inspectDocmanLoginState(page);
  if (authBefore.onLoginPage) {
    throw new Error(
      "Docman login required before verification. " +
        `Current URL: ${authBefore.url}`
    );
  }

  const baseUrl = new URL(page.url()).origin;
  console.log("✔ Docman environment detected:", baseUrl);

  // Navigate to User List using the detected environment.
  await page.goto(`${baseUrl}/Admin/Users/UserList`, {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });

  const authAfterNav = await inspectDocmanLoginState(page);
  if (authAfterNav.onLoginPage) {
    await page
      .screenshot({ path: "verify-users-auth-failed.png", fullPage: true })
      .catch(() => {});
    throw new Error(
      "Docman redirected back to login while opening User List. " +
        `Current URL: ${authAfterNav.url}. Screenshot: verify-users-auth-failed.png`
    );
  }

  await waitForUserListReady(page);

  const filter = page
    .locator(
      [
        "#Filter_Criteria",
        'input[name="Filter.Criteria"]',
        'input[id*="Filter_Criteria"]',
        'input[type="search"]',
      ].join(", ")
    )
    .first();
  await filter.waitFor({ timeout: 60000 });

  const results = [];

  for (const username of usernames) {
    const exactCandidates = await runSearch(page, filter, username);
    const exactMatch = findExactMatch(exactCandidates, username);
    const partialMatches = [];

    if (!exactMatch) {
      const parts = username
        .split(" ")
        .map((p) => p.trim())
        .filter((p) => p.length >= 3)
        .slice(0, 2);

      for (const part of parts) {
        const partCandidates = await runSearch(page, filter, part);
        for (const candidate of partCandidates) {
          const containsPart = candidate
            .toLowerCase()
            .includes(part.toLowerCase());
          if (
            containsPart &&
            !partialMatches.includes(candidate) &&
            !isSameUser(candidate, username)
          ) {
            partialMatches.push(candidate);
          }
        }
        if (partialMatches.length >= 5) break;
      }
    }

    results.push({
      searchedName: username,
      exists: Boolean(exactMatch),
      docmanUsername: exactMatch || null,
      partialMatches: partialMatches.length ? partialMatches : null,
      needsManualReview: !exactMatch && partialMatches.length > 0,
    });
  }

  return results;
}

async function waitForUserListReady(page) {
  await page.waitForSelector("table tbody", { timeout: 60000 });
  await page.waitForSelector(
    [
      "#Filter_Criteria",
      'input[name="Filter.Criteria"]',
      'input[id*="Filter_Criteria"]',
      'input[type="search"]',
    ].join(", "),
    { timeout: 60000 }
  );
}

async function runSearch(page, filter, term) {
  await filter.click({ timeout: 10000 }).catch(() => {});
  await filter.fill("");
  await filter.type(term, { delay: 20 });

  await Promise.allSettled([
    filter.press("Enter"),
    page.waitForLoadState("domcontentloaded", { timeout: 1200 }),
  ]);

  await page.waitForTimeout(120);
  return await readVisibleUsernames(page);
}

async function readVisibleUsernames(page) {
  return await page.$$eval(
    "table tbody tr td:first-child a, table tbody tr td:first-child",
    (cells) => {
      const out = [];
      for (const cell of cells) {
        const value = (cell.textContent || "").trim();
        if (!value) continue;

        const normalized = value.toLowerCase();
        if (
          normalized === "no records found" ||
          normalized === "no matching records found" ||
          normalized === "no data available in table"
        ) {
          continue;
        }

        if (!out.includes(value)) out.push(value);
      }
      return out;
    }
  );
}

async function inspectDocmanLoginState(page) {
  const url = page.url();
  const lowerUrl = url.toLowerCase();

  const loginByUrl =
    lowerUrl.includes("/account/login") || lowerUrl.includes("/account/prelogin");

  const signInHeadingVisible = await page
    .locator("text=/Sign in to Continue/i")
    .first()
    .isVisible({ timeout: 500 })
    .catch(() => false);

  const autoSignInFailedVisible = await page
    .locator("text=/automatic sign-in failed/i")
    .first()
    .isVisible({ timeout: 500 })
    .catch(() => false);

  const orgFieldVisible = await page
    .locator(
      [
        "#OrganisationCode",
        "#OrganizationCode",
        "#OdsCode",
        'input[name="OrganisationCode"]',
        'input[name="OrganizationCode"]',
        'input[name="OdsCode"]',
      ].join(", ")
    )
    .first()
    .isVisible({ timeout: 500 })
    .catch(() => false);

  const userFieldVisible = await page
    .locator(
      [
        "#UserName",
        "#Username",
        'input[name="UserName"]',
        'input[name="Username"]',
      ].join(", ")
    )
    .first()
    .isVisible({ timeout: 500 })
    .catch(() => false);

  const passFieldVisible = await page
    .locator('#Password, input[name="Password"], input[type="password"]')
    .first()
    .isVisible({ timeout: 500 })
    .catch(() => false);

  const onLoginPage =
    loginByUrl ||
    signInHeadingVisible ||
    autoSignInFailedVisible ||
    (orgFieldVisible && (userFieldVisible || passFieldVisible));

  return { onLoginPage, url };
}

function findExactMatch(candidates, target) {
  return candidates.find((candidate) => isSameUser(candidate, target)) || null;
}

function isSameUser(a, b) {
  return normalizeName(a) === normalizeName(b);
}

function normalizeName(value) {
  return (value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

module.exports = verifyDocmanUsers;
