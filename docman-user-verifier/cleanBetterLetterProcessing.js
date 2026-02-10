// cleanBetterLetterProcessing.js
const UUID_REGEX =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

async function cleanBetterLetterProcessing({ page, batchSize = 50, dryRun = false }) {
  try {
    // SOURCE
    const sourceFolder = await promptUntilFolderExists(page, "Enter SOURCE folder name to scan (exact match):");
    if (!sourceFolder) return;

    await loadFilingFolder(page, sourceFolder);

    // SCAN (strong selector)
    const allTitles = await page.$$eval(
      "#document_list li a div strong, #document_list li a strong",
      (els) => els.map((e) => e.innerText.trim()).filter(Boolean)
    );

    if (allTitles.length === 0) {
      throw new Error("Document list is empty — folder may not have loaded or selector mismatch");
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
    const destinationFolder = await promptUntilFolderExists(page, "Enter destination folder name (exact match):");
    if (!destinationFolder) return;

    const proceed = await promptYesNo(`Move ${nonUuidTitles.length} documents to "${destinationFolder}"?`);
    if (!proceed) {
      console.log("Cancelled. No documents were moved.");
      return;
    }

    // MOVE
    await ensureSelectMode(page);

    let remaining = [...nonUuidTitles];
    let batch = 1;

    while (remaining.length) {
      const current = remaining.slice(0, batchSize);
      console.log(`\nBatch ${batch}: moving ${current.length}`);

      await selectDocumentsByTitle(page, current);
      await openChangeFolder(page);
      await moveToFolder(page, destinationFolder);

      remaining = remaining.slice(batchSize);
      batch++;

      await page.waitForTimeout(1000);
      await ensureSelectMode(page);
    }

    console.log("\n✔ All documents moved.");
  } catch (err) {
    console.error("❌ CLEAN FAILED:", err.message);
    await page.screenshot({ path: "clean-failure.png", fullPage: true }).catch(() => {});
    throw err;
  }
}

/* ---------------- folder helpers ---------------- */

async function promptUntilFolderExists(page, promptMsg) {
  while (true) {
    const name = await promptText(promptMsg);

    if (!name) {
      console.log("Cancelled.");
      return null;
    }

    const found = await findFolderLink(page, name);
    if (found) return name;

    console.log(`❌ Folder "${name}" not found.`);
  }
}

// Finds folder by scrolling INSIDE the folder pane until found or end.
async function findFolderLink(page, folderName) {
  const tree = page.locator("#folders_list").first();
  await tree.waitFor({ timeout: 60000 });

  // scroll reset to top
  await tree.evaluate((el) => (el.scrollTop = 0));

  for (let i = 0; i < 80; i++) {
    const link = tree.locator(
      `xpath=.//a[.//span[normalize-space(text())="${folderName}"]]`
    );

    if (await link.count()) return link.first();

    // scroll down
    const didScroll = await tree.evaluate((el) => {
      const before = el.scrollTop;
      el.scrollTop = before + el.clientHeight * 0.9;
      return el.scrollTop !== before;
    });

    if (!didScroll) break; // reached end

    await page.waitForTimeout(80);
  }

  return null;
}

async function loadFilingFolder(page, folderName) {
  console.log(`➡ Opening folder: "${folderName}"`);

  const tree = page.locator("#folders_list").first();
  await tree.waitFor({ timeout: 60000 });

  const link = await findFolderLink(page, folderName);
  if (!link) throw new Error(`Could not load folder "${folderName}" (not found in folder pane)`);

  await link.click({ force: true });

  // Confirm selected
  await page
    .locator(
      `xpath=//*[@id="selectedFolder" and contains(normalize-space(.), "${folderName}")]`
    )
    .waitFor({ timeout: 15000 });

  // Wait for list refresh or empty state
  await page.waitForSelector("#document_list li, .instruction.primary, #document_list", {
    timeout: 15000,
  });

  console.log(`✔ Folder "${folderName}" loaded`);
}

/* ---------------- move helpers ---------------- */

async function ensureSelectMode(page) {
  if (await page.locator('#document_list input[type="checkbox"]').count()) return;

  // Prefer the doc list container’s menu button
  const docList = page.locator("#document_list");
  await docList.waitFor({ timeout: 60000 });

  const menu = docList.locator("button").last();
  await menu.click({ timeout: 60000 });

  const selectMode = page.getByText("Select Mode", { exact: true });
  await selectMode.click({ timeout: 60000 });

  await page.waitForSelector('#document_list input[type="checkbox"]', { timeout: 60000 });
}

async function selectDocumentsByTitle(page, titles) {
  const items = await page.$$("#document_list li");
  for (const item of items) {
    const titleEl = await item.$("a div strong") || await item.$("a strong");
    if (!titleEl) continue;

    const title = (await titleEl.innerText()).trim();
    if (!titles.includes(title)) continue;

    const checkbox = await item.$('input[type="checkbox"]');
    if (checkbox && !(await checkbox.isChecked())) {
      await checkbox.click();
      await page.waitForTimeout(10);
    }
  }
}

async function openChangeFolder(page) {
  await page.getByText("Change Folder", { exact: true }).click({ timeout: 60000 });
  await page.getByText("Change Document Folder", { exact: true }).waitFor({ timeout: 60000 });
}

async function moveToFolder(page, folderName) {
  await page.getByText(folderName, { exact: true }).click({ timeout: 60000 });
  await page.getByRole("button", { name: "Confirm" }).click({ timeout: 60000 });
}

/* ---------------- CLI prompts ---------------- */

function promptYesNo(q) {
  return new Promise((res) => {
    process.stdout.write(`${q} (y/N): `);
    process.stdin.once("data", (d) => res(d.toString().trim().toLowerCase() === "y"));
  });
}

function promptText(q) {
  return new Promise((res) => {
    process.stdout.write(`${q} `);
    process.stdin.once("data", (d) => res(d.toString().trim()));
  });
}

module.exports = cleanBetterLetterProcessing;
