// automation/fetchDocmanCreds.js
const BETTERLETTER_PRACTICES_URL = "https://app.betterletter.ai/admin_panel/practices";

async function fetchDocmanCreds(page, practiceName) {
  const timeoutMs = 60000;

  await page.goto(BETTERLETTER_PRACTICES_URL, {
    waitUntil: "domcontentloaded",
    timeout: timeoutMs,
  });

  // Ensure practices list exists (Phoenix patch links)
  await page.waitForSelector('a[href^="/admin_panel/practices/"]', { timeout: timeoutMs });

  const target = String(practiceName || "").trim();
  if (!target) {
    throw new Error("Practice name is required to fetch Docman credentials.");
  }
  const targetLower = target.toLowerCase();

  // 1) Exact match on visible span text
  let link = page.locator(
    `xpath=//a[starts-with(@href,"/admin_panel/practices/")][.//span[normalize-space(text())="${target}"]]`
  ).first();

  // 2) Case-insensitive contains match
  if ((await link.count()) === 0) {
    link = page.locator(
      `xpath=//a[starts-with(@href,"/admin_panel/practices/")][.//span[contains(translate(normalize-space(text()), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "${targetLower}")]]`
    ).first();
  }

  // 3) Fallback through row text matching
  if ((await link.count()) === 0) {
    const rows = page.locator("table tbody tr");
    const rowCount = await rows.count();
    for (let i = 0; i < rowCount; i += 1) {
      const row = rows.nth(i);
      const rowText = ((await row.innerText().catch(() => "")) || "").toLowerCase();
      if (!rowText.includes(targetLower)) continue;
      const rowLink = row.locator('a[href^="/admin_panel/practices/"]').first();
      if ((await rowLink.count()) > 0) {
        link = rowLink;
        break;
      }
    }
  }

  if ((await link.count()) === 0) {
    // Helpful: show a few practice names visible for debugging
    const names = await page.$$eval(
      'a[href^="/admin_panel/practices/"] span',
      (spans) => spans.map((s) => (s.textContent || "").trim()).filter(Boolean).slice(0, 20)
    );

    throw new Error(
      `Practice "${practiceName}" not found in the visible BetterLetter list.\n` +
      `Try a longer/clearer substring (e.g. "Heathview Medical").\n` +
      `First visible practices:\n- ${names.join("\n- ")}`
    );
  }

  await openPracticeDetailsPage(page, link, timeoutMs);

  await openEhrSettingsTab(page, timeoutMs);

  const odsCode = await readOdsCode(page, timeoutMs);
  const { adminUsername, adminPassword } = await readDocmanInputs(page, timeoutMs);

  if (!odsCode || !adminUsername || !adminPassword) {
    throw new Error("Could not read Docman creds from BetterLetter EHR Settings");
  }

  console.log("✔ Credentials resolved:");
  console.log(`  ODS: ${odsCode}`);
  console.log(`  User: ${adminUsername}`);

  // Best-effort: each practice configures its own Docman folder names here
  // (e.g. "1.For BetterLetter" / "2.Processing by BetterLetter" - naming
  // varies per practice), so read them directly instead of guessing at a
  // fixed convention. Never throws - callers that don't need these (login,
  // verify, etc.) just ignore empty values, and clean-processing falls back
  // to its own folder-name guesses when a field can't be read here.
  const folderNames = await readFolderNames(page, timeoutMs).catch(() => ({
    inputFolder: "",
    processingFolder: "",
    filingFolder: "",
    rejectedFolder: "",
  }));

  return { odsCode, adminUsername, adminPassword, ...folderNames };
}

async function openPracticeDetailsPage(page, link, timeoutMs) {
  const href = String((await link.getAttribute("href").catch(() => "")) || "").trim();
  const expectedPath =
    href && href.startsWith("/admin_panel/practices/") ? href : "";

  if (expectedPath) {
    const absoluteUrl = new URL(expectedPath, BETTERLETTER_PRACTICES_URL).toString();
    await page.goto(absoluteUrl, {
      waitUntil: "domcontentloaded",
      timeout: timeoutMs,
    });
    return;
  }

  // Fallback when href is unavailable/changed: click and wait for patched route.
  await Promise.allSettled([
    page
      .waitForURL(/\/admin_panel\/practices\/[^/?#]+/i, { timeout: timeoutMs })
      .catch(() => null),
    link.click({ timeout: 20000 }),
  ]);
  await page.waitForLoadState("domcontentloaded", { timeout: timeoutMs }).catch(() => {});
}

async function openEhrSettingsTab(page, timeoutMs) {
  const tabSelectors = [
    '[data-test-id="tab-ehr_settings"]',
    '[phx-value-tab="ehr_settings"]',
    'a[href*="ehr_settings"]',
    'button:has-text("EHR Settings")',
    'a:has-text("EHR Settings")',
  ];

  const tabSelectorList = tabSelectors.join(", ");
  const tab = page.locator(tabSelectorList).first();

  await tab.waitFor({ state: "attached", timeout: timeoutMs }).catch(async () => {
    await captureFetchCredsDebug(page, "ehr-tab-missing");
    throw new Error(
      `EHR Settings tab not found on practice page. URL: ${page.url()}. ` +
      "Saved debug files: fetch-docman-creds-ehr-tab-missing.png/.html"
    );
  });

  // Try direct click paths first.
  try {
    await tab.scrollIntoViewIfNeeded().catch(() => {});
    await tab.click({ timeout: 10000 });
  } catch (_) {
    try {
      await tab.click({ timeout: 10000, force: true });
    } catch (_) {
      // Last resort: DOM-level click on known selectors.
      const clicked = await page
        .evaluate((selectors) => {
          for (const selector of selectors) {
            const el = document.querySelector(selector);
            if (!el) continue;
            el.scrollIntoView({ block: "center" });
            el.click();
            return true;
          }
          return false;
        }, tabSelectors)
        .catch(() => false);
      if (!clicked) {
        await captureFetchCredsDebug(page, "ehr-tab-click-failed");
        throw new Error(
          `Failed to open EHR Settings tab. URL: ${page.url()}. ` +
          "Saved debug files: fetch-docman-creds-ehr-tab-click-failed.png/.html"
        );
      }
    }
  }

  await page.waitForTimeout(300);
}

async function readOdsCode(page, timeoutMs) {
  const odsCandidates = [
    "span.text-white.bg-subtle",
    "span.bg-subtle",
    '[data-test-id*="ods" i]',
  ];

  for (const selector of odsCandidates) {
    const el = page.locator(selector).first();
    const count = await el.count().catch(() => 0);
    if (!count) continue;

    await el.waitFor({ state: "attached", timeout: Math.min(timeoutMs, 20000) }).catch(() => {});
    const text = (await el.innerText().catch(() => "")).trim();
    const ods = extractOdsFromText(text);
    if (ods) return ods;
  }

  // Fallback: scan page text for an ODS-like token.
  const bodyText = (await page.locator("body").innerText().catch(() => "")) || "";
  const bodyOds = extractOdsFromText(bodyText);
  if (bodyOds) return bodyOds;

  // Final fallback: ODS often appears in practice path.
  const urlOds = extractOdsFromText(page.url());
  if (urlOds) return urlOds;

  return "";
}

async function readDocmanInputs(page, timeoutMs) {
  const userInput = firstAvailableLocator(page, [
    '#ehr_settings\\[docman\\]\\[username\\]',
    'input[name="ehr_settings[docman][username]"]',
    'input[name*="[docman][username]"]',
  ]);
  const passInput = firstAvailableLocator(page, [
    '#ehr_settings\\[docman\\]\\[password\\]',
    'input[name="ehr_settings[docman][password]"]',
    'input[name*="[docman][password]"]',
  ]);

  await userInput.waitFor({ state: "attached", timeout: timeoutMs }).catch(async () => {
    await captureFetchCredsDebug(page, "docman-username-missing");
    throw new Error(
      `Docman username input not found in EHR Settings. URL: ${page.url()}. ` +
      "Saved debug files: fetch-docman-creds-docman-username-missing.png/.html"
    );
  });

  await passInput.waitFor({ state: "attached", timeout: timeoutMs }).catch(async () => {
    await captureFetchCredsDebug(page, "docman-password-missing");
    throw new Error(
      `Docman password input not found in EHR Settings. URL: ${page.url()}. ` +
      "Saved debug files: fetch-docman-creds-docman-password-missing.png/.html"
    );
  });

  const adminUsername = (await userInput.inputValue().catch(() => "")).trim();
  const adminPassword = await passInput.inputValue().catch(() => "");

  return { adminUsername, adminPassword };
}

async function readFolderNames(page, timeoutMs) {
  const [inputFolder, processingFolder, filingFolder, rejectedFolder] = await Promise.all([
    readFieldByLabel(page, timeoutMs, "Input Folder", [
      'input[name*="[input_folder]"]',
      'input[name*="[folder_input]"]',
    ]),
    readFieldByLabel(page, timeoutMs, "Processing Folder", [
      'input[name*="[processing_folder]"]',
      'input[name*="[folder_processing]"]',
    ]),
    readFieldByLabel(page, timeoutMs, "Filing Folder", [
      'input[name*="[filing_folder]"]',
      'input[name*="[folder_filing]"]',
    ]),
    readFieldByLabel(page, timeoutMs, "Rejected Folder", [
      'input[name*="[rejected_folder]"]',
      'input[name*="[folder_rejected]"]',
    ]),
  ]);
  return { inputFolder, processingFolder, filingFolder, rejectedFolder };
}

// Reads a folder-name value by its visible section label (e.g. "PROCESSING
// FOLDER" on the practice's EHR Settings page), rather than a hardcoded form
// field name - each practice sets its own Docman folder names here, so
// there's no fixed convention to select by. Tries a same-container <input>
// first in case it's an editable field, then falls back to reading whatever
// plain text sits next to the label, since this section may just display
// the configured name rather than let it be edited from this page.
async function readFieldByLabel(page, timeoutMs, labelText, nameSelectorGuesses = []) {
  if (nameSelectorGuesses.length) {
    const guessed = page.locator(nameSelectorGuesses.join(", ")).first();
    if ((await guessed.count().catch(() => 0)) > 0) {
      const value = (await guessed.inputValue().catch(() => "")).trim();
      if (value) return value;
    }
  }

  const escapedLabel = labelText.toLowerCase();
  const label = page.locator(
    `xpath=//*[contains(translate(normalize-space(text()), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "${escapedLabel}") and string-length(normalize-space(text())) < 60]`
  ).first();

  const found = await label.waitFor({ state: "attached", timeout: Math.min(timeoutMs, 15000) })
    .then(() => true)
    .catch(() => false);
  if (!found) return "";

  const container = label.locator("xpath=ancestor::*[self::div or self::section][1]").first();
  if ((await container.count().catch(() => 0)) === 0) return "";

  const input = container.locator("input").first();
  if ((await input.count().catch(() => 0)) > 0) {
    const value = (await input.inputValue().catch(() => "")).trim();
    if (value) return value;
  }

  const containerText = (await container.innerText().catch(() => "")) || "";
  const labelOwnText = ((await label.innerText().catch(() => "")) || "").trim().toLowerCase();
  const remainder = containerText
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line && line.toLowerCase() !== labelOwnText)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();

  return remainder;
}

function firstAvailableLocator(page, selectors) {
  return page.locator(selectors.join(", ")).first();
}

function extractOdsFromText(text) {
  const value = String(text || "");
  const match = value.match(/\b[A-Z]\d{5}\b/);
  return match ? match[0] : "";
}

async function captureFetchCredsDebug(page, suffix) {
  const safe = String(suffix || "debug").replace(/[^a-z0-9_-]+/gi, "-");
  const screenshotPath = `fetch-docman-creds-${safe}.png`;
  const htmlPath = `fetch-docman-creds-${safe}.html`;

  await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
  const html = await page.content().catch(() => "");
  if (html) {
    const fs = require("fs");
    fs.writeFileSync(htmlPath, html, "utf8");
  }
}

module.exports = { fetchDocmanCreds };
