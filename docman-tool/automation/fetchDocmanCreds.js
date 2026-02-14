// automation/fetchDocmanCreds.js
async function fetchDocmanCreds(page, practiceName) {
  await page.goto("https://app.betterletter.ai/admin_panel/practices", {
    waitUntil: "domcontentloaded",
    timeout: 60000,
  });

  // Ensure practices list exists (Phoenix patch links)
  await page.waitForSelector('a[href^="/admin_panel/practices/"]', { timeout: 60000 });

  const target = practiceName.trim();
  const targetLower = target.toLowerCase();

  // 1) Exact match on the visible <span> practice name
  let link = page.locator(
    `xpath=//a[starts-with(@href,"/admin_panel/practices/")][.//span[normalize-space(text())="${target}"]]`
  ).first();

  // 2) Fallback: contains match (case-insensitive)
  if ((await link.count()) === 0) {
    link = page.locator(
      `xpath=//a[starts-with(@href,"/admin_panel/practices/")][.//span[contains(translate(normalize-space(text()), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "${targetLower}")]]`
    ).first();
  }

  if ((await link.count()) === 0) {
    // Helpful: show a few practice names visible for debugging
    const names = await page.$$eval(
      'a[href^="/admin_panel/practices/"] span',
      (spans) => spans.map((s) => s.textContent.trim()).filter(Boolean).slice(0, 20)
    );

    throw new Error(
      `Practice "${practiceName}" not found in the visible BetterLetter list.\n` +
      `Try a longer/clearer substring (e.g. "Heathview Medical").\n` +
      `First visible practices:\n- ${names.join("\n- ")}`
    );
  }

  await link.click();
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
