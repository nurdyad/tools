// automation/fetchDocmanCreds.js
async function fetchDocmanCreds(page, practiceName) {
  // Go to practices list
  await page.goto("https://app.betterletter.ai/admin_panel/practices", {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });

  // Try exact match first, then fallback to contains
  const exact = page.locator("a", { hasText: practiceName }).filter({
    has: page.locator(`text="${practiceName}"`),
  });

  let practiceLink = exact.first();

  if (!(await practiceLink.count())) {
    practiceLink = page.locator("a", { hasText: practiceName }).first();
  }

  if (!(await practiceLink.count())) {
    throw new Error(`Practice "${practiceName}" not found in BetterLetter list`);
  }

  await practiceLink.click();
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
