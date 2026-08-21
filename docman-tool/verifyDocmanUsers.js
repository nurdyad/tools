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

  let groupNames = [];
  try {
    groupNames = await readAllRecipientGroupNames(page, baseUrl);
    console.log(`✔ Loaded ${groupNames.length} Docman recipient group(s).`);
  } catch (error) {
    console.log(
      `⚠ Could not load Docman recipient groups, continuing with users only: ${error.message}`
    );
  }

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

  const filter = await waitForUserListReady(page);

  const results = [];

  for (const username of usernames) {
    const exactCandidates = await runSearch(page, filter, username);
    let exactMatch = findBestResolvedMatch(exactCandidates, username);
    let matchType = exactMatch ? "user" : null;
    const partialMatches = [];

    if (!exactMatch) {
      addRelevantPartialMatches(partialMatches, exactCandidates, username);
      const searchTerms = buildFallbackSearchTerms(username);

      for (const part of searchTerms) {
        const partCandidates = await runSearch(page, filter, part);
        const resolvedFromPart = findBestResolvedMatch(partCandidates, username);
        if (resolvedFromPart) {
          exactMatch = resolvedFromPart;
          matchType = "user";
          break;
        }

        addRelevantPartialMatches(partialMatches, partCandidates, username);
        if (partialMatches.length >= 5) break;
      }
    }

    if (!exactMatch && groupNames.length) {
      const exactGroupMatch = findBestResolvedMatch(groupNames, username);
      if (exactGroupMatch) {
        exactMatch = exactGroupMatch;
        matchType = "group";
      } else if (partialMatches.length < 5) {
        addRelevantPartialMatches(partialMatches, groupNames, username);
      }
    }

    results.push({
      searchedName: username,
      exists: Boolean(exactMatch),
      docmanUsername: exactMatch || null,
      matchType,
      partialMatches: partialMatches.length ? partialMatches : null,
      needsManualReview: !exactMatch && partialMatches.length > 0,
    });
  }

  return results;
}

async function readAllRecipientGroupNames(page, baseUrl) {
  await page.goto(`${baseUrl}/Admin/RecipientGroups/RecipientGroupList`, {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });

  const authState = await inspectDocmanLoginState(page);
  if (authState.onLoginPage) {
    throw new Error(
      `Docman redirected to login while opening Recipient Groups. Current URL: ${authState.url}`
    );
  }

  await page.waitForSelector("table", { timeout: 60000 });

  const names = new Set();
  const addCurrentPageRows = async () => {
    for (const name of await readRecipientGroupTableRows(page)) {
      names.add(name);
    }
  };

  await addCurrentPageRows();

  const pageHrefs = await page.evaluate(() =>
    Array.from(
      document.querySelectorAll('a.dropdown-item[href*="RecipientGroups/PageMove"]')
    ).map((a) => a.getAttribute("href"))
  );

  // The pagination links live inside a collapsed dropdown, so they aren't
  // "visible" to Playwright's normal click - dispatch a real click on the
  // anchor directly (a plain page.goto to the same href does not reliably
  // advance Docman's server-side paging state).
  for (const href of pageHrefs) {
    const clicked = await page.evaluate((targetHref) => {
      const link = document.querySelector(`a.dropdown-item[href="${targetHref}"]`);
      if (!link) return false;
      link.click();
      return true;
    }, href);
    if (!clicked) continue;

    await page.waitForLoadState("domcontentloaded", { timeout: 15000 }).catch(() => {});
    await page.waitForTimeout(300);
    await addCurrentPageRows();
  }

  return [...names];
}

async function readRecipientGroupTableRows(page) {
  return await page.evaluate(() => {
    const tables = Array.from(document.querySelectorAll("table"));
    const groupTable = tables.find((table) =>
      /description/i.test((table.querySelector("thead") || {}).textContent || "")
    );
    if (!groupTable) return [];

    return Array.from(groupTable.querySelectorAll("tbody tr"))
      .map((row) => {
        const firstCell = row.querySelector("td");
        return firstCell ? firstCell.textContent.trim() : "";
      })
      .filter(Boolean);
  });
}

async function waitForUserListReady(page) {
  await page.waitForSelector("table tbody, table", { timeout: 60000 });
  return await waitForVisibleFilter(page, 60000);
}

async function runSearch(page, filter, term) {
  const baselineSnapshot = await readUserListSnapshot(page);

  await filter.click({ timeout: 10000 }).catch(() => {});
  await filter.fill("").catch(() => {});
  await filter.type(term, { delay: 20 }).catch(async () => {
    await filter.fill(term).catch(() => {});
  });
  await filter.evaluate((element) => {
    element.dispatchEvent(new Event("input", { bubbles: true }));
    element.dispatchEvent(new Event("change", { bubbles: true }));
  }).catch(() => {});

  await Promise.allSettled([
    filter.press("Enter"),
    page.waitForLoadState("networkidle", { timeout: 1200 }),
  ]);

  await waitForResultsToSettle(page, filter, term, baselineSnapshot);
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

function findBestResolvedMatch(candidates, target) {
  const exact = candidates.find((candidate) => isSameUser(candidate, target, { stripTitles: false }));
  if (exact) return exact;

  const normalizedMatches = candidates.filter((candidate) =>
    isSameUser(candidate, target, { stripTitles: true })
  );

  return normalizedMatches.length === 1 ? normalizedMatches[0] : null;
}

function isSameUser(a, b, options = {}) {
  return normalizeName(a, options) === normalizeName(b, options);
}

function normalizeName(value, options = {}) {
  const stripTitles = options.stripTitles !== false;
  let normalized = String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ");

  if (stripTitles) {
    normalized = normalized.replace(/\b(mr|mrs|miss|ms|dr|prof|professor|sir|lady)\b/g, " ");
  }

  return normalized.replace(/\s+/g, " ").trim();
}

function buildFallbackSearchTerms(username) {
  const raw = String(username || "").trim();
  const stripped = normalizeName(raw, { stripTitles: true });
  const terms = [];

  if (stripped && stripped !== normalizeName(raw, { stripTitles: false })) {
    terms.push(stripped);
  }

  const tokens = stripped
    .split(" ")
    .map((token) => token.trim())
    .filter((token) => token.length >= 3);

  for (const token of tokens) {
    if (!terms.includes(token)) terms.push(token);
  }

  if (tokens.length >= 2) {
    const firstLast = `${tokens[0]} ${tokens[tokens.length - 1]}`.trim();
    if (firstLast && !terms.includes(firstLast)) terms.push(firstLast);
  }

  return terms.slice(0, 5);
}

function addRelevantPartialMatches(partialMatches, candidates, username) {
  for (const candidate of candidates) {
    if (partialMatches.length >= 5) break;
    if (partialMatches.includes(candidate)) continue;
    if (isRelevantPartialMatch(candidate, username)) {
      partialMatches.push(candidate);
    }
  }
}

function isRelevantPartialMatch(candidate, username) {
  const candidateRaw = normalizeName(candidate, { stripTitles: false });
  const usernameRaw = normalizeName(username, { stripTitles: false });
  if (!candidateRaw || candidateRaw === usernameRaw) return false;

  const candidateNormalized = normalizeName(candidate, { stripTitles: true });
  const usernameNormalized = normalizeName(username, { stripTitles: true });
  if (!candidateNormalized || !usernameNormalized) return false;

  if (candidateNormalized === usernameNormalized) return true;

  const usernameTokens = usernameNormalized.split(" ").filter(Boolean);
  const candidateTokens = candidateNormalized.split(" ").filter(Boolean);
  if (!usernameTokens.length || !candidateTokens.length) return false;

  if (usernameTokens.every((token) => candidateTokens.includes(token))) {
    return true;
  }

  const { exactMatches, fuzzyMatches, matchedCount } = countMatchedNameTokens(
    usernameTokens,
    candidateTokens
  );

  if (!exactMatches) return false;

  if (matchedCount >= usernameTokens.length) {
    return true;
  }

  if (usernameTokens.length >= 2 && candidateTokens.length >= 2) {
    const firstTokenMatches = areSimilarNameTokens(usernameTokens[0], candidateTokens[0]);
    const lastTokenMatches = areSimilarNameTokens(
      usernameTokens[usernameTokens.length - 1],
      candidateTokens[candidateTokens.length - 1]
    );

    if (firstTokenMatches && lastTokenMatches) {
      return matchedCount >= Math.max(2, usernameTokens.length - 1) || fuzzyMatches > 0;
    }
  }

  return false;
}

function countMatchedNameTokens(usernameTokens, candidateTokens) {
  const remaining = [...candidateTokens];
  let exactMatches = 0;
  let fuzzyMatches = 0;

  for (const usernameToken of usernameTokens) {
    const exactIndex = remaining.indexOf(usernameToken);
    if (exactIndex !== -1) {
      exactMatches += 1;
      remaining.splice(exactIndex, 1);
      continue;
    }

    const fuzzyIndex = remaining.findIndex((candidateToken) =>
      areSimilarNameTokens(usernameToken, candidateToken)
    );
    if (fuzzyIndex !== -1) {
      fuzzyMatches += 1;
      remaining.splice(fuzzyIndex, 1);
    }
  }

  return {
    exactMatches,
    fuzzyMatches,
    matchedCount: exactMatches + fuzzyMatches,
  };
}

function areSimilarNameTokens(a, b) {
  if (a === b) return true;

  const left = String(a || "").trim();
  const right = String(b || "").trim();
  if (!left || !right) return false;

  const shorter = left.length <= right.length ? left : right;
  const longer = left.length <= right.length ? right : left;

  // Short names are very often a nickname that's a literal prefix of the full
  // name (Sam/Samantha, Ben/Benjamin, Al/Albert), so allow substring
  // containment down to 2 characters instead of requiring 4+.
  if (shorter.length >= 2 && longer.includes(shorter)) {
    return true;
  }

  if (shorter.length < 4) return false;

  const distance = getLevenshteinDistance(left, right);
  const maxLength = Math.max(left.length, right.length);

  if (maxLength <= 5) return distance <= 1;
  if (maxLength <= 8) return distance <= 2;
  return distance <= 3;
}

function getLevenshteinDistance(a, b) {
  const rows = a.length + 1;
  const cols = b.length + 1;
  const distances = Array.from({ length: rows }, (_, row) => {
    const values = new Array(cols).fill(0);
    values[0] = row;
    return values;
  });

  for (let col = 0; col < cols; col += 1) {
    distances[0][col] = col;
  }

  for (let row = 1; row < rows; row += 1) {
    for (let col = 1; col < cols; col += 1) {
      const substitutionCost = a[row - 1] === b[col - 1] ? 0 : 1;
      distances[row][col] = Math.min(
        distances[row - 1][col] + 1,
        distances[row][col - 1] + 1,
        distances[row - 1][col - 1] + substitutionCost
      );
    }
  }

  return distances[rows - 1][cols - 1];
}

async function waitForVisibleFilter(page, timeoutMs) {
  const selectors = [
    "#Filter_Criteria",
    'input[name="Filter.Criteria"]',
    'input[id*="Filter_Criteria"]',
    'xpath=//label[contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"hide inactive")]/following::input[not(@type="checkbox") and not(@type="radio") and not(@type="hidden")][1]',
    'xpath=//h1[contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"user list")]/following::input[not(@type="checkbox") and not(@type="radio") and not(@type="hidden")][1]',
    'input[placeholder*="search" i]',
    'input[type="search"]',
  ];

  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    for (const selector of selectors) {
      const locator = page.locator(selector).first();
      const visible = await locator.isVisible({ timeout: 150 }).catch(() => false);
      if (visible) return locator;
    }
    await page.waitForTimeout(200);
  }

  throw new Error("User list filter input not visible");
}

async function waitForResultsToSettle(page, filter, term, baselineSnapshot, timeoutMs = 6000) {
  const startedAt = Date.now();
  const normalizedTerm = normalizeName(term, { stripTitles: false });
  let lastSnapshot = "";
  let stableCount = 0;
  let sawChange = false;

  while (Date.now() - startedAt < timeoutMs) {
    const inputValue = normalizeName(await filter.inputValue().catch(() => ""), {
      stripTitles: false,
    });
    const snapshot = await readUserListSnapshot(page);
    if (snapshot !== baselineSnapshot) {
      sawChange = true;
    }

    const minWaitMs = sawChange ? 300 : 900;
    if (inputValue === normalizedTerm && Date.now() - startedAt >= minWaitMs) {
      if (snapshot === lastSnapshot) {
        stableCount += 1;
      } else {
        stableCount = 0;
      }

      if (stableCount >= 2) {
        return;
      }
    }

    lastSnapshot = snapshot;
    await page.waitForTimeout(200);
  }
}

async function readUserListSnapshot(page) {
  const users = await readVisibleUsernames(page);
  const tableText = await page.locator("table tbody").innerText().catch(() => "");
  return JSON.stringify({
    users,
    tableText: String(tableText || "").replace(/\s+/g, " ").trim(),
  });
}

module.exports = verifyDocmanUsers;
