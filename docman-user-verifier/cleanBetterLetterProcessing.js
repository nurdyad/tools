const fs = require("fs");
const path = require("path");
const inquirer = require("inquirer").default;

const UUID_REGEX =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

async function cleanBetterLetterProcessing({ page, batchSize = 50, dryRun = false }) {
  try {
    const scope = await resolveFilingScope(page);
    console.log(`ℹ CLEAN scope resolved: ${scope === page ? "main page" : "iframe"}`);
    await loadFilingScreen(scope);

    // SOURCE
    const sourceFolder = await promptUntilFolderLoads(
      scope,
      "Enter SOURCE folder name to scan (exact match):"
    );
    if (!sourceFolder) return;

    // SCAN (strong selector)
    const allTitles = await scope.$$eval(
      "#document_list li a div strong, #document_list li a strong",
      (els) => els.map((e) => e.innerText.trim()).filter(Boolean)
    );

    if (allTitles.length === 0) {
      throw new Error(
        "Document list is empty — folder may not have loaded or selector mismatch"
      );
    }

    const nonUuidTitles = allTitles.filter((t) => !UUID_REGEX.test(t));

    console.log(`\n📄 Documents detected: ${allTitles.length}`);
    console.log(`Found ${nonUuidTitles.length} NON-UUID documents.`);

    if (!nonUuidTitles.length) {
      console.log("Nothing to move.");
      return;
    }

    console.log("\nExamples:");
    nonUuidTitles.slice(0, 10).forEach((t) => console.log(" -", t));

    if (dryRun) {
      console.log("\n🟡 DRY RUN — no changes made.");
      return;
    }

    // DESTINATION
    const destinationFolder = await promptUntilFolderExists(
      scope,
      "Enter destination folder name (exact match):"
    );
    if (!destinationFolder) return;

    const proceed = await promptYesNo(
      `Move ${nonUuidTitles.length} documents to "${destinationFolder}"?`
    );
    if (!proceed) {
      console.log("Cancelled. No documents were moved.");
      return;
    }

    // MOVE
    await ensureSelectMode(scope);

    let remaining = [...nonUuidTitles];
    let batch = 1;

    while (remaining.length) {
      const current = remaining.slice(0, batchSize);
      console.log(`\nBatch ${batch}: moving ${current.length}`);

      await selectDocumentsByTitle(scope, current);
      await openChangeFolder(scope);
      await moveToFolder(scope, destinationFolder);

      remaining = remaining.slice(batchSize);
      batch++;

      await waitForTimeout(scope, 700);
      await ensureSelectMode(scope);
    }

    console.log("\n✔ All documents moved.");
  } catch (err) {
    console.error("❌ CLEAN FAILED:", err.message);
    await page.screenshot({ path: "clean-failure.png", fullPage: true }).catch(() => {});
    throw err;
  }
}

/* ---------------- scope helpers ---------------- */

async function resolveFilingScope(page) {
  if (await isFilingScopeReady(page)) return page;

  for (const frame of page.frames()) {
    if (frame === page.mainFrame()) continue;
    if (await isFilingScopeReady(frame)) return frame;
  }

  await page
    .screenshot({ path: "clean-filing-scope-not-found.png", fullPage: true })
    .catch(() => {});
  throw new Error(
    "Could not find Docman Filing folder pane (#folders_list/#folders) in page or iframes. " +
      "Screenshot: clean-filing-scope-not-found.png"
  );
}

async function isFilingScopeReady(scope) {
  const tree = scope.locator("#folders_list, #folders").first();
  if (await tree.isVisible({ timeout: 800 }).catch(() => false)) return true;
  if ((await tree.count().catch(() => 0)) > 0) return true;

  const docs = scope.locator("#document_list").first();
  if (await docs.isVisible({ timeout: 800 }).catch(() => false)) return true;
  if ((await docs.count().catch(() => 0)) > 0) return true;

  return false;
}

function getScopePage(scope) {
  return typeof scope.page === "function" ? scope.page() : scope;
}

async function waitForTimeout(scope, ms) {
  await getScopePage(scope).waitForTimeout(ms);
}

/* ---------------- folder helpers ---------------- */

async function promptUntilFolderLoads(scope, promptMsg) {
  while (true) {
    const name = await promptText(promptMsg);

    if (!name) {
      console.log("Cancelled.");
      return null;
    }

    try {
      console.log(`🔎 Trying to load source folder: "${name}"`);
      await withTimeout(
        loadFilingFolder(scope, name),
        30000,
        `Timed out while loading folder "${name}"`
      );
      return name;
    } catch (err) {
      console.log(`❌ Folder "${name}" could not be loaded. ${err?.message || ""}`);
    }
  }
}

async function promptUntilFolderExists(scope, promptMsg) {
  while (true) {
    const name = await promptText(promptMsg);

    if (!name) {
      console.log("Cancelled.");
      return null;
    }

    const found = await findFolderLink(scope, name);
    if (found) return name;

    console.log(`❌ Folder "${name}" not found.`);
  }
}

async function loadFilingScreen(scope) {
  const allDocs = scope.locator("span.all-docs-count").first();
  if ((await allDocs.count().catch(() => 0)) > 0) {
    await allDocs.click().catch(() => {});
  }

  await scope.waitForSelector("#folders, #folders_list", { timeout: 10000 });
  await getScopePage(scope).waitForLoadState("domcontentloaded").catch(() => {});
}

// Finds folder by scrolling INSIDE the folder pane until found or end.
async function findFolderLink(scope, folderName) {
  await withTimeout(loadFilingScreen(scope), 10000, "Filing screen not ready");

  const tree = scope.locator("#folders_list, #folders").first();
  await tree.waitFor({ state: "attached", timeout: 20000 });

  const target = folderName.trim();
  const targetLower = target.toLowerCase();
  const exactPattern = new RegExp(`^\\s*${escapeRegExp(target)}\\s*$`, "i");
  const containsPattern = new RegExp(escapeRegExp(target), "i");

  // Scroll to top
  await tree.evaluate((el) => {
    el.scrollTop = 0;
  }).catch(() => {});

  async function toClickable(locator) {
    const clickable = locator
      .locator(
        "xpath=ancestor-or-self::a[1] | ancestor-or-self::button[1] | ancestor-or-self::li[1]"
      )
      .first();
    if ((await clickable.count().catch(() => 0)) > 0) return await firstVisible(clickable);
    return await firstVisible(locator);
  }

  async function fromXPath(xpath) {
    const loc = tree.locator(`xpath=${xpath}`);
    if ((await loc.count().catch(() => 0)) === 0) return null;
    return await toClickable(loc);
  }

  async function tryFind() {
    // 1) Exact text on a folder span (closest to your Python implementation)
    const exactSpan = await fromXPath(
      `.//span[normalize-space(text())=${toXPathLiteral(target)}] | ` +
      `.//a[normalize-space(text())=${toXPathLiteral(target)}]`
    );
    if (exactSpan) return exactSpan;

    // 2) Case-insensitive contains on span text
    const containsSpan = await fromXPath(
      `.//span[contains(translate(normalize-space(text()), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), ${toXPathLiteral(
        targetLower
      )})] | ` +
      `.//a[contains(translate(normalize-space(text()), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), ${toXPathLiteral(
        targetLower
      )})]`
    );
    if (containsSpan) return containsSpan;

    // 3) Playwright text matching fallbacks
    const exact = tree.getByText(exactPattern).first();
    if ((await exact.count().catch(() => 0)) > 0) return await toClickable(exact);

    const contains = tree.getByText(containsPattern).first();
    if ((await contains.count().catch(() => 0)) > 0) return await toClickable(contains);

    return null;
  }

  for (let i = 0; i < 90; i++) {
    const hit = await withTimeout(
      tryFind(),
      1500,
      `Folder lookup iteration timed out for "${target}"`
    ).catch(() => null);
    if (hit) {
      await hit.scrollIntoViewIfNeeded().catch(() => {});
      return hit;
    }

    if (i > 0 && i % 10 === 0) {
      console.log(`…still searching for folder "${target}" (pass ${i})`);
    }

    const didScroll = await tree
      .evaluate((el) => {
        const before = el.scrollTop;
        el.scrollTop = before + el.clientHeight * 0.9;
        return el.scrollTop !== before;
      })
      .catch(() => false);

    if (!didScroll) break;
    await waitForTimeout(scope, 50);
  }

  return null;
}

async function loadFilingFolder(scope, folderName) {
  let lastError = null;
  let lastDebug = null;

  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      console.log(`🔁 Folder load attempt ${attempt}/3: "${folderName}"`);
      await withTimeout(
        loadFilingFolderOnce(scope, folderName),
        15000,
        `Timed out loading folder "${folderName}" on attempt ${attempt}`
      );

      await scope.waitForSelector(
        'xpath=//ul[@id="document_list"] | //div[contains(@class,"instruction") and contains(@class,"primary")]',
        { timeout: 5000 }
      );
      return;
    } catch (error) {
      lastError = error;
      lastDebug = await withTimeout(
        saveFolderDebugArtifacts({
          scope,
          folderName,
          attempt,
          note: error?.message || "unknown error",
        }),
        4000,
        "Timed out while writing clean debug artifacts"
      ).catch(() => null);
      if (lastDebug) {
        console.log(
          `🧪 Debug saved: ${lastDebug.jsonFile} and ${lastDebug.screenshotFile}`
        );
      }

      if (attempt < 3) {
        console.log(
          `⚠ Retry ${attempt}/3 while loading folder "${folderName}" (${error.message})`
        );
        continue;
      }
      throw new Error(
        `Could not load folder "${folderName}" after 3 attempts. ` +
          `Last error: ${lastError?.message || "unknown"}. ` +
          (lastDebug
            ? `Debug files: ${lastDebug.jsonFile}, ${lastDebug.screenshotFile}`
            : "")
      );
    }
  }
}

async function loadFilingFolderOnce(scope, folderName) {
  console.log(`➡ Opening folder: "${folderName}"`);
  await loadFilingScreen(scope);

  const link = await findFolderLink(scope, folderName);
  if (!link) {
    throw new Error(`Could not load folder "${folderName}" (not found in folder pane)`);
  }

  await link.click({ force: true, timeout: 5000 });

  await scope.waitForSelector(
    `xpath=//*[@id='selectedFolder' and contains(normalize-space(.), ${toXPathLiteral(
      folderName
    )})]`,
    { timeout: 5000 }
  );
  await scope.waitForSelector("#document_list, #document_list li, .instruction.primary", {
    timeout: 5000,
  });
  await waitForTimeout(scope, 1000);

  console.log(`✔ Folder "${folderName}" loaded`);
}

/* ---------------- move helpers ---------------- */

async function ensureSelectMode(scope) {
  if ((await scope.locator('#document_list input[type="checkbox"]').count()) > 0) return;

  const docList = scope.locator("#document_list").first();
  await docList.waitFor({ timeout: 60000 });

  const menu = docList.locator("button").last();
  await menu.click({ timeout: 60000 });

  const selectMode = scope.locator("text=/^Select Mode$/i").first();
  await selectMode.waitFor({ timeout: 60000 });
  await selectMode.click({ timeout: 60000 });

  await scope.waitForSelector('#document_list input[type="checkbox"]', {
    timeout: 60000,
  });
}

async function selectDocumentsByTitle(scope, titles) {
  const items = await scope.$$("#document_list li");
  let selected = 0;

  for (const item of items) {
    const titleEl = (await item.$("a div strong")) || (await item.$("a strong"));
    if (!titleEl) continue;

    const title = (await titleEl.innerText()).trim();
    if (!titles.includes(title)) continue;

    const checkbox = await item.$('input[type="checkbox"]');
    if (checkbox && !(await checkbox.isChecked())) {
      await checkbox.click();
      selected++;
      await waitForTimeout(scope, 10);
    }
  }

  if (!selected) {
    console.log("⚠ No matching visible documents were selected in this batch.");
  }
}

async function openChangeFolder(scope) {
  const byId = scope.locator("a#action_changefolder").first();
  if ((await byId.count().catch(() => 0)) > 0) {
    await byId.click({ timeout: 60000 });
  } else {
    await scope.locator("text=/^Change Folder$/i").first().click({ timeout: 60000 });
  }

  await scope
    .locator("text=/^Change Document Folder$/i")
    .first()
    .waitFor({ timeout: 60000 });
}

async function moveToFolder(scope, folderName) {
  const targetByDataName = scope
    .locator(
      `xpath=//ul[@id="folderselection"]//a[@data-name=${toXPathLiteral(
        folderName
      )}] | //ul[@id="folderselection"]//li/a[contains(normalize-space(.), ${toXPathLiteral(folderName)})]`
    )
    .first();
  const targetFallback = scope
    .locator(`xpath=//*[normalize-space(text())=${toXPathLiteral(folderName)}]`)
    .first();

  if ((await targetByDataName.count().catch(() => 0)) > 0) {
    await targetByDataName.click({ timeout: 60000 });
  } else {
    await targetFallback.waitFor({ timeout: 60000 });
    await targetFallback.click({ timeout: 60000 });
  }

  const confirmById = scope.locator("input#change_folder_confirm").first();
  if ((await confirmById.count().catch(() => 0)) > 0) {
    await confirmById.click({ timeout: 60000 });
  } else {
    await scope.locator('button:has-text("Confirm")').first().click({ timeout: 60000 });
  }
}

/* ---------------- misc helpers ---------------- */

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function toXPathLiteral(value) {
  if (!value.includes("'")) return `'${value}'`;
  if (!value.includes('"')) return `"${value}"`;
  const parts = value.split("'");
  return `concat('${parts.join(`', "'", '`)}')`;
}

async function firstVisible(locator, maxItems = 25) {
  const count = await locator.count().catch(() => 0);
  if (!count) return locator.first();

  const limit = Math.min(count, Math.min(maxItems, 10));
  for (let i = 0; i < limit; i++) {
    const item = locator.nth(i);
    const visible = await item.isVisible({ timeout: 30 }).catch(() => false);
    if (visible) return item;
  }

  return locator.first();
}

async function saveFolderDebugArtifacts({ scope, folderName, attempt, note }) {
  try {
    const page = getScopePage(scope);
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const safeFolder = sanitizeFileName(folderName || "unknown-folder");
    const base = `clean-folder-debug-${safeFolder}-${stamp}-attempt-${attempt}`;
    const jsonFile = `${base}.json`;
    const screenshotFile = `${base}.png`;

    const tree = scope.locator("#folders_list, #folders").first();
    const folderSamples = await tree
      .locator("span, a, li")
      .evaluateAll((nodes) => {
        const uniq = [];
        for (const n of nodes) {
          const text = (n.textContent || "").replace(/\s+/g, " ").trim();
          if (!text) continue;
          if (text.length > 120) continue;
          if (!uniq.includes(text)) uniq.push(text);
          if (uniq.length >= 250) break;
        }
        return uniq;
      })
      .catch(() => []);

    const treeHtmlSnippet = await tree
      .evaluate((el) => (el?.innerHTML || "").slice(0, 20000))
      .catch(() => "");

    const info = {
      createdAt: new Date().toISOString(),
      folderName,
      attempt,
      note,
      pageUrl: page.url(),
      scopeType: scope === page ? "page" : "frame",
      scopeUrl: typeof scope.url === "function" ? scope.url() : page.url(),
      frameUrls: page.frames().map((f) => f.url()),
      selectors: {
        foldersCount: await scope.locator("#folders, #folders_list").count().catch(() => 0),
        documentListCount: await scope.locator("#document_list").count().catch(() => 0),
        allDocsCountBadge: await scope.locator("span.all-docs-count").count().catch(() => 0),
        selectedFolderCount: await scope.locator("#selectedFolder").count().catch(() => 0),
      },
      folderSamples,
      treeHtmlSnippet,
    };

    const jsonPath = path.join(process.cwd(), jsonFile);
    fs.writeFileSync(jsonPath, JSON.stringify(info, null, 2), "utf8");
    await page.screenshot({ path: screenshotFile, fullPage: true }).catch(() => {});

    return { jsonFile, screenshotFile };
  } catch (_) {
    return null;
  }
}

function sanitizeFileName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60);
}

function withTimeout(promise, ms, message) {
  let timeoutId = null;
  const timeoutPromise = new Promise((_, reject) => {
    timeoutId = setTimeout(() => reject(new Error(message || "Operation timed out")), ms);
  });

  return Promise.race([promise, timeoutPromise]).finally(() => {
    if (timeoutId) clearTimeout(timeoutId);
  });
}

/* ---------------- CLI prompts ---------------- */

function promptYesNo(q) {
  return inquirer
    .prompt([
      {
        type: "confirm",
        name: "value",
        message: q,
        default: false,
      },
    ])
    .then((ans) => Boolean(ans.value));
}

function promptText(q) {
  return inquirer
    .prompt([
      {
        type: "input",
        name: "value",
        message: q,
      },
    ])
    .then((ans) => (ans.value || "").trim());
}

module.exports = cleanBetterLetterProcessing;
