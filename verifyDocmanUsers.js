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

      // Clear + type search
      await filter.fill("");
      await filter.type(username, { delay: 50 });
      await filter.press("Enter");

      // Wait for Docman to rebuild the table
      await page.waitForFunction(() => {
        const rows = document.querySelectorAll("table tbody tr");
        return rows.length > 0;
      }, { timeout: 15000 });

      let foundUsername = null;

      const firstRowLink = page.locator(
        "table tbody tr:first-child td:first-child a"
      );

      if (await firstRowLink.count()) {
        foundUsername = (await firstRowLink.innerText())
          .replace(/^\*/, "")
          .trim();
      }

      results.push({
        searchedName: username,
        exists: Boolean(foundUsername),
        docmanUsername: foundUsername
      });
    }


    return results;
  } finally {
    await browser.close();
  }
}

module.exports = verifyDocmanUsers;
