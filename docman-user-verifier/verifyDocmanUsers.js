const { chromium } = require("playwright");

async function verifyDocmanUsers({
  odsCode,
  adminUsername,
  adminPassword,
  usernames
}) {
  const browser = await chromium.launch({
    headless: false
  });

  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    console.log(`Checking Docman users for ODS: ${odsCode}`);

    /* -------- Login to Docman -------- */

    await page.goto(
      "https://production.docman.thirdparty.nhs.uk",
      { waitUntil: "domcontentloaded", timeout: 60000 }
    );

    await page.fill("#OdsCode", odsCode);
    await page.fill("#Username", adminUsername);
    await page.fill("#Password", adminPassword);

    await Promise.all([
      page.click('button[type="submit"]'),
      page.waitForLoadState("domcontentloaded")
    ]);


    // Navigate via direct URL AFTER home is ready
    await page.goto(
      "https://production.docman.thirdparty.nhs.uk/Admin/Users/UserList",
      { waitUntil: "domcontentloaded" }
    );

    // Wait until the users table actually exists
    await page.waitForSelector("table tbody", { timeout: 60000 });

    const results = [];

    for (const username of usernames) {
      const filter = page.locator("#Filter_Criteria");

      async function runSearch(term) {
        await filter.fill("");
        await filter.type(term, { delay: 50 });
        await filter.press("Enter");

        try {
          await page.waitForFunction(() => {
            const rows = document.querySelectorAll("table tbody tr");
            return rows.length > 0;
          }, { timeout: 6000 });

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
        needsManualReview: !exactMatch && partialMatches.length > 0
      });
    }


    return results;
  } finally {
    await browser.close();
  }
}

module.exports = verifyDocmanUsers;
