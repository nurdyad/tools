// automation/fetchDocmanCreds.js
async function fetchDocmanCreds(page, practiceName) {
  await page.goto("https://app.betterletter.ai/admin_panel/practices", {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });

  // If BetterLetter returned Unauthorized or login screen, fail clearly
  const bodyText = ((await page.textContent("body").catch(() => "")) || "").trim();
  if (/unauthorized/i.test(bodyText)) {
    throw new Error(
      "BetterLetter returned Unauthorized (HTTP Basic Auth failed). " +
      "Check your Basic Auth username/password in run.js."
    );
  }
  if (page.url().includes("/users/log_in") || page.url().includes("/login")) {
    throw new Error(
      "BetterLetter admin panel requires login. Please login in the opened browser first."
    );
  }

  // Wait for practices table/list to appear
  await page.waitForSelector("table, #practices, body", { timeout: 60000 });

  // Prefer clicking the Practice Name link within the table row.
  // We'll do case-insensitive contains matching and also handle partial names.
  const target = practiceName.trim().toLowerCase();

  const rows = page.locator("table tbody tr");
  const rowCount = await rows.count();

  if (rowCount === 0) {
    throw new Error("No practices rows found. The page may not have loaded correctly.");
  }

  let matchedRowIndex = -1;
  let matchedPracticeText = null;

  for (let i = 0; i < rowCount; i++) {
    const row = rows.nth(i);
    const rowText = ((await row.innerText().catch(() => "")) || "").toLowerCase();

    if (rowText.includes(target)) {
      matchedRowIndex = i;
      matchedPracticeText = rowText;
      break;
    }
  }

  if (matchedRowIndex === -1) {
    // Helpful: show a few examples for debugging
    const sample = [];
    const max = Math.min(8, rowCount);
    for (let i = 0; i < max; i++) {
      const t = ((await rows.nth(i).innerText().catch(() => "")) || "").trim();
      sample.push(t.split("\n")[0]);
    }

    throw new Error(
      `Practice "${practiceName}" not found in BetterLetter list.\n` +
      `Try a partial name (e.g. "Heathview") or the exact table name.\n` +
      `Example rows:\n- ${sample.join("\n- ")}`
    );
  }

  const matchedRow = rows.nth(matchedRowIndex);

  // Click the first link inside the practice row (usually practice name)
  const rowLink = matchedRow.locator("a").first();
  if (!(await rowLink.count())) {
    throw new Error(
      `Found the practice row but no clickable link inside it.\nRow text:\n${matchedPracticeText}`
    );
  }

  await rowLink.click();
  await page.waitForLoadState("domcontentloaded");

  // Click EHR Settings tab
  const ehrTab = page.locator('[data-test-id="tab-ehr_settings"]');
  await ehrTab.waitFor({ timeout: 60000 });
  await ehrTab.click();

  // ODS code
  const odsEl = page.locator("span.bg-subtle").first();
  await odsEl.waitFor({ timeout: 60000 });
  const odsCode = (await odsEl.innerText()).trim();

  // Docman username & password
  const userInput = page.locator('#ehr_settings\\[docman\\]\\[username\\]');
  const passInput = page.locator('#ehr_settings\\[docman\\]\\[password\\]');

  await userInput.waitFor({ timeout: 60000 });
  await passInput.waitFor({ timeout: 60000 });

  const adminUsername = (await userInput.inputValue()).trim();
  const adminPassword = await passInput.inputValue();

  if (!odsCode || !adminUsername || !adminPassword) {
    throw new Error("Could not read Docman creds from BetterLetter EHR Settings");
  }

  console.log("✔ Credentials resolved:");
  console.log(`  ODS: ${odsCode}`);
  console.log(`  User: ${adminUsername}`);

  return { odsCode, adminUsername, adminPassword };
}

module.exports = { fetchDocmanCreds };
