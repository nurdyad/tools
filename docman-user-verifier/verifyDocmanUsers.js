// verifyDocmanUsers.js
async function verifyDocmanUsers({ page, usernames }) {
  if (!page) {
    throw new Error("verifyDocmanUsers requires an existing Playwright page.");
  }

  console.log("Checking Docman users in existing session.");

  // Ensure we are logged in by detecting login screen
  if (page.url().includes("/Account/Login")) {
    throw new Error(
      "Docman login required. Please complete login in the opened browser window."
    );
  }

  const baseUrl = new URL(page.url()).origin;
  console.log("✔ Docman environment detected:", baseUrl);

  // Navigate to User List using the detected environment
  await page.goto(`${baseUrl}/Admin/Users/UserList`, {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });

  // Confirm the users table is present
  await page.waitForSelector("table tbody tr", { timeout: 60000 });

  const results = [];

  for (const username of usernames) {
    const filter = page.locator("#Filter_Criteria");

    async function runSearch(term) {
      await filter.fill("");
      await filter.type(term, { delay: 50 });
      await filter.press("Enter");

      try {
        await page.waitForFunction(
          () => {
            const rows = document.querySelectorAll("table tbody tr");
            return rows.length > 0;
          },
          { timeout: 6000 }
        );

        const link = page.locator(
          "table tbody tr:first-child td:first-child a"
        );

        if (await link.count()) {
          return await link.innerText();
        }
      } catch {
        return null;
      }

      return null;
    }

    // 1️⃣ Full name search first
    let exactMatch = await runSearch(username);
    let partialMatches = [];

    if (!exactMatch) {
      const parts = username.split(" ").filter(Boolean);

      for (const part of parts) {
        const match = await runSearch(part);
        if (match && !partialMatches.includes(match)) {
          partialMatches.push(match);
        }
      }
    }

    results.push({
      searchedName: username,
      exists: Boolean(exactMatch),
      docmanUsername: exactMatch,
      partialMatches: partialMatches.length ? partialMatches : null,
      needsManualReview: !exactMatch && partialMatches.length > 0,
    });
  }

  return results;
}

module.exports = verifyDocmanUsers;
