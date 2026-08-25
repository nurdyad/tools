const fs = require("fs");
const path = require("path");
const inquirer = require("inquirer").default;
const { withRetry } = require("./automation/retry");
const { classifyError } = require("./automation/runLogger");

const UUID_REGEX =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
// Fallback guesses only - practices can name these folders almost anything
// (Grove Medical Centre uses "1.For BetterLetter" / "2.Processing by
// BetterLetter" / "3.Filing by BetterLetter", a third convention besides
// these two), so the authoritative source is each practice's own EHR
// Settings folder-name fields (see fetchDocmanCreds.js), passed in via
// inputs.sourceFolder/destinationFolder and used before ever reaching this
// list. This only matters when that lookup comes back empty.
const CLEAN_FOLDER_GROUPS = {
  processing: [
    "BetterLetter: Processing",
    "zz BL Processing. Do not touch",
  ],
  filing: [
    "BetterLetter: Filing",
    "zz BL Filing. Do not touch",
  ],
  input: [
    "BetterLetter: Input",
    "zz BL Input. Do not touch",
  ],
};
const FILING_DATE_REGEX = /^\d{4}-\d{2}-\d{2}$/;
const FILING_DOCUMENT_ID_REGEX = /^[A-Za-z0-9][A-Za-z0-9-]{2,}$/;
const FILING_TITLE_SEGMENT_REGEX = /^[A-Za-z0-9][A-Za-z0-9 &'()\/.,:-]*$/;

function normalizeRetryPolicy(retryPolicy = {}) {
  const attempts = Number.isInteger(retryPolicy?.attempts) ? retryPolicy.attempts : 1;
  const baseDelayMs = Number.isFinite(retryPolicy?.baseDelayMs)
    ? Math.max(0, Math.floor(retryPolicy.baseDelayMs))
    : 0;
  const maxDelayMs = Number.isFinite(retryPolicy?.maxDelayMs)
    ? Math.max(baseDelayMs, Math.floor(retryPolicy.maxDelayMs))
    : baseDelayMs;

  return {
    attempts: Math.max(1, attempts),
    baseDelayMs,
    maxDelayMs,
  };
}

function normalizeFolderPickerOptions(folderPicker = {}) {
  return {
    enabled: folderPicker?.enabled !== false,
    maxSuggestions: Number.isInteger(folderPicker?.maxSuggestions)
      ? Math.min(Math.max(folderPicker.maxSuggestions, 3), 25)
      : 10,
    maxScrollPasses: Number.isInteger(folderPicker?.maxScrollPasses)
      ? Math.min(Math.max(folderPicker.maxScrollPasses, 10), 400)
      : 90,
  };
}

async function cleanBetterLetterProcessing({
  page,
  cleanType = "",
  batchSize = 50,
  dryRun = false,
  defaults = {},
  inputs = {},
  folderPicker = {},
  retryPolicy = {},
  logger = null,
}) {
  const resolvedRetry = normalizeRetryPolicy(retryPolicy);
  const pickerOptions = normalizeFolderPickerOptions(folderPicker);
  const sourceDefault = String(defaults?.sourceFolder || "").trim();
  const destinationDefault = String(defaults?.destinationFolder || "").trim();
  const sourceInput = String(inputs?.sourceFolder || "").trim();
  const destinationInput = String(inputs?.destinationFolder || "").trim();
  const autoConfirmMove = Boolean(inputs?.autoConfirmMove);
  const nonInteractive = Boolean(inputs?.nonInteractive);
  const normalizedCleanType = normalizeCleanType(cleanType);
  const cleanProfile = getCleanProfile(normalizedCleanType);

  try {
    const scope = await resolveFilingScope(page);
    logger?.event("clean_scope_resolved", {
      scopeType: scope === page ? "main_page" : "iframe",
      cleanType: normalizedCleanType || "manual",
      dryRun,
      batchSize,
    });

    console.log(`ℹ CLEAN scope resolved: ${scope === page ? "main page" : "iframe"}`);
    await loadFilingScreen(scope);

    const availableFolders = cleanProfile
      ? await collectFolderNames(scope, pickerOptions).catch(() => [])
      : [];

    // SOURCE
    let sourceFolder = sourceInput;
    if (sourceFolder) {
      console.log(`✔ Enter SOURCE folder name to scan (exact match): ${sourceFolder}`);
      console.log(`🔎 Trying to load source folder: "${sourceFolder}"`);
      await withTimeout(
        loadFilingFolder(scope, sourceFolder),
        30000,
        `Timed out while loading folder "${sourceFolder}"`
      );
    } else if (cleanProfile) {
      sourceFolder = await resolveFirstExistingFolder(
        scope,
        cleanProfile.sourceFolderCandidates,
        availableFolders
      );
      if (!sourceFolder) {
        console.log(
          `ℹ Folders actually found in the tree (${availableFolders.length}): ${
            availableFolders.length ? availableFolders.join(" | ") : "(none - tree may not have loaded)"
          }`
        );
        throw new Error(
          `Could not find a ${cleanProfile.label} source folder. Tried: ${cleanProfile.sourceFolderCandidates.join(
            " | "
          )}. Folders actually found (${availableFolders.length}): ${
            availableFolders.length ? availableFolders.slice(0, 15).join(" | ") : "none - tree may not have loaded"
          }`
        );
      }
      console.log(`✔ CLEAN source folder auto-resolved: ${sourceFolder}`);
      await withTimeout(
        loadFilingFolder(scope, sourceFolder),
        30000,
        `Timed out while loading folder "${sourceFolder}"`
      );
    } else {
      sourceFolder = await promptUntilFolderLoads(
        scope,
        "Enter SOURCE folder name to scan:",
        {
          defaultFolder: sourceDefault,
          pickerOptions,
        }
      );
      if (!sourceFolder) return;
    }

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

    const titlesToMove = allTitles.filter((title) =>
      cleanProfile ? cleanProfile.shouldMoveTitle(title) : !UUID_REGEX.test(title)
    );

    logger?.event("clean_scan_complete", {
      cleanType: normalizedCleanType || "manual",
      sourceFolder,
      totalDocuments: allTitles.length,
      matchedDocuments: titlesToMove.length,
    });

    console.log(`\n📄 Documents detected: ${allTitles.length}`);
    console.log(`Found ${titlesToMove.length} ${cleanProfile?.matchLabel || "NON-UUID documents"}.`);

    if (!titlesToMove.length) {
      console.log("Nothing to move.");
      return;
    }

    console.log("\nExamples:");
    titlesToMove.slice(0, 10).forEach((t) => console.log(" -", t));

    if (dryRun) {
      console.log("\n🟡 DRY RUN — no changes made.");
      return;
    }

    // DESTINATION
    let destinationFolder = destinationInput;
    if (destinationFolder) {
      console.log(`✔ Enter destination folder name (exact match): ${destinationFolder}`);
      const found = await findFolderLinkWithOptions(scope, destinationFolder, {
        prepare: false,
      });
      if (!found) {
        throw new Error(`Folder "${destinationFolder}" not found.`);
      }
    } else if (cleanProfile) {
      destinationFolder = await resolveFirstExistingFolder(
        scope,
        cleanProfile.destinationFolderCandidates,
        availableFolders
      );
      if (!destinationFolder) {
        console.log(
          `ℹ Folders actually found in the tree (${availableFolders.length}): ${
            availableFolders.length ? availableFolders.join(" | ") : "(none - tree may not have loaded)"
          }`
        );
        throw new Error(
          `Could not find an input folder. Tried: ${cleanProfile.destinationFolderCandidates.join(
            " | "
          )}. Folders actually found (${availableFolders.length}): ${
            availableFolders.length ? availableFolders.slice(0, 15).join(" | ") : "none - tree may not have loaded"
          }`
        );
      }
      console.log(`✔ CLEAN destination folder auto-resolved: ${destinationFolder}`);
    } else {
      destinationFolder = await promptUntilFolderExists(
        scope,
        "Enter destination folder name:",
        {
          nonDisruptive: true,
          defaultFolder: destinationDefault,
          pickerOptions,
        }
      );
      if (!destinationFolder) return;
    }

    let proceed = autoConfirmMove;
    if (proceed) {
      console.log(`✔ Move ${titlesToMove.length} documents to "${destinationFolder}"? Yes`);
    } else {
      proceed = await promptYesNo(
        `Move ${titlesToMove.length} documents to "${destinationFolder}"?`
      );
    }

    if (!proceed) {
      if (nonInteractive) {
        throw new Error("Move confirmation required in non-interactive mode. Use --yes.");
      }
      console.log("Cancelled. No documents were moved.");
      return;
    }

    // Destination lookup can change the UI context in some tenants.
    // Re-open source folder to guarantee we select the intended documents.
    await withTimeout(
      loadFilingFolder(scope, sourceFolder),
      20000,
      `Timed out reloading source folder "${sourceFolder}" before move`
    );

    // MOVE
    await ensureSelectMode(scope);

    let remaining = [...titlesToMove];
    let batch = 1;

    while (remaining.length) {
      const current = remaining.slice(0, batchSize);
      const startedAt = Date.now();
      console.log(`\nBatch ${batch}: moving ${current.length}`);

      logger?.event("clean_batch_start", {
        batch,
        inBatch: current.length,
        remainingBefore: remaining.length,
        destinationFolder,
      });

      await withRetry(
        async () => {
          await dismissTransientBlockingModals(scope, "before selecting documents");
          await selectDocumentsByTitle(scope, current);
          await dismissTransientBlockingModals(scope, "after selecting documents");
          await openChangeFolder(scope);
          await dismissTransientBlockingModals(scope, "before choosing destination folder");
          await moveToFolder(scope, destinationFolder);
          await dismissTransientBlockingModals(scope, "after confirming move");
        },
        {
          ...resolvedRetry,
          label: `clean batch ${batch}`,
          onRetry: ({ nextAttempt, attempts, delayMs, error }) => {
            console.log(
              `↻ Batch ${batch} failed (${error?.message || "unknown error"}). ` +
                `retry ${nextAttempt}/${attempts} in ${delayMs}ms`
            );
            logger?.event("clean_batch_retry", {
              batch,
              nextAttempt,
              attempts,
              delayMs,
              errorType: classifyError(error),
              errorMessage: error?.message || "unknown error",
            });
          },
        }
      );

      const durationMs = Date.now() - startedAt;
      logger?.event("clean_batch_end", {
        batch,
        moved: current.length,
        durationMs,
      });

      remaining = remaining.slice(batchSize);
      batch++;

      if (remaining.length) {
        await waitForTimeout(scope, 700);
        await ensureSelectMode(scope);
      }
    }

    console.log("\n✔ All documents moved.");
  } catch (err) {
    logger?.event("clean_failed", {
      errorType: classifyError(err),
      errorMessage: err?.message || String(err),
    });
    console.error("❌ CLEAN FAILED:", err.message);
    await page.screenshot({ path: "clean-failure.png", fullPage: true }).catch(() => {});
    throw err;
  }
}

function normalizeCleanType(cleanTypeInput) {
  const cleanType = String(cleanTypeInput || "").trim().toLowerCase();
  if (
    cleanType === "processing" ||
    cleanType === "process" ||
    cleanType === "proc" ||
    cleanType === "processing-folder"
  ) {
    return "processing";
  }
  if (
    cleanType === "filing" ||
    cleanType === "file" ||
    cleanType === "fiiling" ||
    cleanType === "filing-folder"
  ) {
    return "filing";
  }
  return "";
}

function getCleanProfile(cleanType) {
  if (cleanType === "processing") {
    return {
      label: "processing",
      sourceFolderCandidates: CLEAN_FOLDER_GROUPS.processing,
      destinationFolderCandidates: CLEAN_FOLDER_GROUPS.input,
      matchLabel: "NON-UUID documents",
      shouldMoveTitle: (title) => !UUID_REGEX.test(String(title || "").trim()),
    };
  }

  if (cleanType === "filing") {
    return {
      label: "filing",
      sourceFolderCandidates: CLEAN_FOLDER_GROUPS.filing,
      destinationFolderCandidates: CLEAN_FOLDER_GROUPS.input,
      matchLabel: "filing titles that do not match the BetterLetter pattern",
      shouldMoveTitle: (title) => !matchesExpectedFilingTitle(title),
    };
  }

  return null;
}

function matchesExpectedFilingTitle(title) {
  const normalized = String(title || "").trim();
  if (!normalized) return false;

  const parts = normalized
    .split("_")
    .map((part) => part.trim())
    .filter(Boolean);

  if (parts.length < 5) {
    return false;
  }

  const datePart = parts[parts.length - 2];
  const documentIdPart = parts[parts.length - 1];
  if (!FILING_DATE_REGEX.test(datePart)) {
    return false;
  }
  if (!FILING_DOCUMENT_ID_REGEX.test(documentIdPart)) {
    return false;
  }

  return parts
    .slice(0, -2)
    .every((part) => FILING_TITLE_SEGMENT_REGEX.test(part));
}

async function resolveFirstExistingFolder(scope, candidates, availableFolders = []) {
  for (const candidate of candidates) {
    const exact = findExactFolderMatch(availableFolders, candidate);
    if (exact) {
      return exact;
    }
  }

  for (const candidate of candidates) {
    const found = await findFolderLinkWithOptions(scope, candidate, {
      prepare: false,
    });
    if (found) {
      return candidate;
    }
  }

  return "";
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

async function resolveActionScope(scope, { requireDocumentList = false } = {}) {
  const page = getScopePage(scope);
  const candidates = [];
  const addCandidate = (candidate) => {
    if (candidate && !candidates.includes(candidate)) candidates.push(candidate);
  };

  addCandidate(scope);
  addCandidate(page.mainFrame());
  for (const frame of page.frames()) addCandidate(frame);

  for (const candidate of candidates) {
    const documentListCount = await candidate
      .locator("#document_list")
      .count()
      .catch(() => 0);
    const checkboxCount = await candidate
      .locator('#document_list input[type="checkbox"]')
      .count()
      .catch(() => 0);

    if (requireDocumentList) {
      if (documentListCount > 0 || checkboxCount > 0) return candidate;
      continue;
    }

    const actionButtonCount = await candidate
      .locator(
        "a#action_selectmode, button#action_selectmode, a#action_changefolder, button#action_changefolder"
      )
      .count()
      .catch(() => 0);
    const actionTextCount = await candidate
      .locator("text=/^Select Mode$/i, text=/^Change Folder$/i")
      .count()
      .catch(() => 0);
    const folderSelectionCount = await candidate
      .locator("#folderselection")
      .count()
      .catch(() => 0);

    if (
      documentListCount > 0 ||
      checkboxCount > 0 ||
      actionButtonCount > 0 ||
      actionTextCount > 0 ||
      folderSelectionCount > 0
    ) {
      return candidate;
    }
  }

  return scope;
}

/* ---------------- folder helpers ---------------- */

async function promptUntilFolderLoads(scope, promptMsg, options = {}) {
  const { defaultFolder = "", pickerOptions = { enabled: true } } = options;
  const availableFolders = await collectFolderNames(scope, pickerOptions).catch(() => []);

  if (pickerOptions.enabled && availableFolders.length) {
    console.log(`ℹ Folder picker indexed ${availableFolders.length} folder(s).`);
  }

  while (true) {
    const name = await promptFolderChoice(promptMsg, {
      defaultFolder,
      availableFolders,
      pickerOptions,
    });

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

async function promptUntilFolderExists(scope, promptMsg, options = {}) {
  const {
    nonDisruptive = false,
    defaultFolder = "",
    pickerOptions = { enabled: true },
  } = options;

  const availableFolders = await collectFolderNames(scope, pickerOptions).catch(() => []);

  if (pickerOptions.enabled && availableFolders.length) {
    console.log(`ℹ Folder picker indexed ${availableFolders.length} folder(s).`);
  }

  while (true) {
    const name = await promptFolderChoice(promptMsg, {
      defaultFolder,
      availableFolders,
      pickerOptions,
    });

    if (!name) {
      console.log("Cancelled.");
      return null;
    }

    const found = await findFolderLinkWithOptions(scope, name, {
      prepare: !nonDisruptive,
    });
    if (found) return name;

    console.log(`❌ Folder "${name}" not found.`);
  }
}

async function promptFolderChoice(promptMsg, options = {}) {
  const {
    defaultFolder = "",
    availableFolders = [],
    pickerOptions = { enabled: true, maxSuggestions: 10 },
  } = options;

  if (!pickerOptions.enabled || !availableFolders.length) {
    return await promptText(promptMsg, defaultFolder);
  }

  let query = await promptText(`${promptMsg} (type a folder name or part of it):`, defaultFolder);

  while (true) {
    if (!query) return "";

    const exact = findExactFolderMatch(availableFolders, query);
    if (exact) return exact;

    const suggestions = getFolderSuggestions(
      availableFolders,
      query,
      pickerOptions.maxSuggestions
    );

    if (!suggestions.length) {
      console.log(`No folder matches found for "${query}".`);
      const useTyped = await promptYesNo(`Use exactly "${query}" anyway?`);
      if (useTyped) return query;

      query = await promptText("Try another folder search (or ENTER to cancel):");
      continue;
    }

    console.log("\nTop folder matches:");
    console.log(` 0. Use exactly: ${query}`);
    suggestions.forEach((name, index) => {
      console.log(` ${index + 1}. ${name}`);
    });

    const pick = await promptText(
      `Choose number 0-${suggestions.length}, or type a new search:`
    );

    if (!pick) return "";

    const asNumber = Number(pick);
    if (Number.isInteger(asNumber) && asNumber === 0) {
      return query;
    }
    if (Number.isInteger(asNumber) && asNumber >= 1 && asNumber <= suggestions.length) {
      return suggestions[asNumber - 1];
    }

    const typedExact = findExactFolderMatch(availableFolders, pick);
    if (typedExact) return typedExact;

    query = pick;
  }
}

async function collectFolderNames(scope, pickerOptions = {}) {
  if (pickerOptions.enabled === false) return [];

  await withTimeout(loadFilingScreen(scope), 10000, "Filing screen not ready for folder picker");

  const tree = scope.locator("#folders_list, #folders").first();
  await tree.waitFor({ state: "attached", timeout: 20000 });

  const folderMap = new Map();
  const maxScrollPasses = Number.isInteger(pickerOptions.maxScrollPasses)
    ? pickerOptions.maxScrollPasses
    : 90;

  await tree
    .evaluate((el) => {
      el.scrollTop = 0;
    })
    .catch(() => {});

  for (let i = 0; i < maxScrollPasses; i++) {
    const labels = await tree
      .locator("li, a, span")
      .evaluateAll((nodes) => {
        const output = [];
        for (const node of nodes) {
          const text = (node.textContent || "").replace(/\s+/g, " ").trim();
          if (!text) continue;
          if (text.length < 2) continue;
          if (text.length > 100) continue;
          output.push(text);
        }
        return output;
      })
      .catch(() => []);

    for (const raw of labels) {
      const normalized = normalizeFolderLabel(raw);
      if (!normalized) continue;
      const key = normalized.toLowerCase();
      if (!folderMap.has(key)) folderMap.set(key, normalized);
    }

    const didScroll = await tree
      .evaluate((el) => {
        const before = el.scrollTop;
        el.scrollTop = before + el.clientHeight * 0.9;
        return el.scrollTop !== before;
      })
      .catch(() => false);

    if (!didScroll) break;
    await waitForTimeout(scope, 30);
  }

  await tree
    .evaluate((el) => {
      el.scrollTop = 0;
    })
    .catch(() => {});

  return Array.from(folderMap.values()).sort((a, b) => a.localeCompare(b));
}

function normalizeFolderLabel(text) {
  const compact = String(text || "").replace(/\s+/g, " ").trim();
  if (!compact) return "";

  const noCount = compact.replace(/\s+\d+$/, "").trim();
  return noCount || compact;
}

function findExactFolderMatch(folders, query) {
  const normalizedQuery = String(query || "").trim().toLowerCase();
  if (!normalizedQuery) return null;

  return (
    folders.find((name) => String(name || "").trim().toLowerCase() === normalizedQuery) || null
  );
}

function getFolderSuggestions(folders, query, maxSuggestions = 10) {
  const normalizedQuery = String(query || "").trim().toLowerCase();
  const max = Number.isInteger(maxSuggestions) ? maxSuggestions : 10;
  const scored = [];

  for (const name of folders) {
    const score = scoreFolderCandidate(name, normalizedQuery);
    if (score < 0) continue;
    scored.push({ name, score });
  }

  scored.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    return a.name.localeCompare(b.name);
  });

  return scored.slice(0, max).map((item) => item.name);
}

function scoreFolderCandidate(name, query) {
  const candidate = String(name || "").toLowerCase();
  if (!candidate) return -1;
  if (!query) return 1;
  if (candidate === query) return 10000;
  if (candidate.startsWith(query)) return 8000 - Math.abs(candidate.length - query.length);

  const idx = candidate.indexOf(query);
  if (idx >= 0) return 6000 - idx;

  if (isSubsequence(query, candidate)) {
    return 3000 - (candidate.length - query.length);
  }

  return -1;
}

function isSubsequence(needle, haystack) {
  if (!needle) return true;
  let i = 0;
  for (let j = 0; j < haystack.length; j++) {
    if (haystack[j] === needle[i]) i += 1;
    if (i >= needle.length) return true;
  }
  return false;
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
  return await findFolderLinkWithOptions(scope, folderName, { prepare: true });
}

async function findFolderLinkWithOptions(scope, folderName, options = {}) {
  const { prepare = true } = options;

  if (prepare) {
    await withTimeout(loadFilingScreen(scope), 10000, "Filing screen not ready");
  }

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

  const link = await findFolderLinkWithOptions(scope, folderName, { prepare: true });
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
  scope = await resolveActionScope(scope, { requireDocumentList: true });

  if ((await scope.locator("#document_list").count().catch(() => 0)) === 0) {
    await scope.waitForSelector("#document_list", { timeout: 4000 }).catch(() => {});
  }

  if ((await scope.locator("#document_list").count().catch(() => 0)) === 0) {
    await saveMoveDebugArtifacts(scope, "document-list-not-found-before-select-mode");
    throw new Error(
      "Document list not found while enabling Select Mode. " +
        "Debug files: clean-move-debug-document-list-not-found-before-select-mode.json/png"
    );
  }

  if (await isSelectModeEnabled(scope)) return;

  const docList = scope.locator("#document_list").first();
  await docList.waitFor({ state: "attached", timeout: 5000 });

  // Strategy 1: direct select mode action button/link.
  const directSelect = scope.locator("a#action_selectmode, button#action_selectmode").first();
  if ((await directSelect.count().catch(() => 0)) > 0) {
    await directSelect.click({ timeout: 5000 }).catch(() => {});
  }
  if (await isSelectModeEnabled(scope)) return;

  // Strategy 2: visible "Select Mode" action.
  const selectModeText = scope.locator("text=/^Select Mode$/i").first();
  const selectModeVisible = await selectModeText.isVisible({ timeout: 500 }).catch(() => false);
  if (selectModeVisible) {
    await selectModeText.click({ timeout: 5000 }).catch(() => {});
  }
  if (await isSelectModeEnabled(scope)) return;

  // Strategy 2b: open document list overflow ("...") then choose select mode.
  const menuOpened = await openDocumentListOverflowMenu(scope);
  if (menuOpened) {
    const clicked = await tryEnableSelectModeFromMenu(scope);
    if (clicked) {
      await waitForTimeout(scope, 400);
      if (await isSelectModeEnabled(scope)) return;
      console.log(
        "⚠ Select Mode checkbox/action was clicked. Continuing (checked state not detectable in DOM)."
      );
      return;
    }
  }
  if (await isSelectModeEnabled(scope)) return;

  // Strategy 3: open common menu candidates then click "Select Mode".
  const menuCandidates = [
    "#document_list button",
    '#document_list [role="button"]',
    '#document_list a[aria-haspopup="true"]',
    "#document_list .dropdown-toggle",
    "#document_list .btn",
  ];

  for (const selector of menuCandidates) {
    const candidates = scope.locator(selector);
    const count = Math.min(await candidates.count().catch(() => 0), 6);
    for (let i = 0; i < count; i++) {
      const menu = candidates.nth(i);
      const visible = await menu.isVisible({ timeout: 200 }).catch(() => false);
      if (!visible) continue;

      await menu.click().catch(() => {});
      const selectMode = scope.locator("text=/^Select Mode$/i").first();
      const visibleSelect = await selectMode.isVisible({ timeout: 500 }).catch(() => false);
      if (visibleSelect) {
        const clicked = await selectMode.click({ timeout: 5000 }).then(() => true).catch(() => false);
        if (clicked) {
          await waitForTimeout(scope, 120);
          if (await isSelectModeEnabled(scope)) return;
          console.log(
            "⚠ Select Mode action clicked from menu candidate. Continuing without checkbox-state confirmation."
          );
          return;
        }
      }

      if (await isSelectModeEnabled(scope)) return;
    }
  }

  await saveMoveDebugArtifacts(scope, "select-mode-not-available");
  throw new Error(
    "Could not enable Select Mode in Document List. " +
      "Debug files: clean-move-debug-select-mode-not-available.json/png"
  );
}

async function selectDocumentsByTitle(scope, titles) {
  scope = await resolveActionScope(scope, { requireDocumentList: true });

  const items = await scope.$$("#document_list li");
  let selected = 0;

  for (const item of items) {
    const titleEl = (await item.$("a div strong")) || (await item.$("a strong"));
    if (!titleEl) continue;

    const title = (await titleEl.innerText()).trim();
    if (!titles.includes(title)) continue;

    const checkbox = await item.$('input[type="checkbox"]');
    if (checkbox) {
      if (!(await checkbox.isChecked())) {
        await checkbox.click();
      }
      selected++;
      await waitForTimeout(scope, 10);
      continue;
    }

    // Some Docman variants use row-click selection in Select Mode (no per-row checkbox).
    const rowClickable =
      (await item.$("a")) ||
      (await item.$("div")) ||
      item;
    await rowClickable.click().catch(() => {});
    selected++;
    await waitForTimeout(scope, 20);
  }

  if (!selected) {
    console.log("⚠ No matching visible documents were selected in this batch.");
  }
}

async function openChangeFolder(scope) {
  scope = await resolveActionScope(scope);
  for (let attempt = 1; attempt <= 2; attempt++) {
    await dismissTransientBlockingModals(scope, "before opening change folder");

    const byId = scope.locator("a#action_changefolder").first();
    const byIdVisible = await byId.isVisible({ timeout: 300 }).catch(() => false);
    if (byIdVisible || (await byId.count().catch(() => 0)) > 0) {
      await byId.click({ timeout: 60000 }).catch(() => {});
    } else if (
      await scope
        .locator("text=/^Change Folder$/i")
        .first()
        .isVisible({ timeout: 400 })
        .catch(() => false)
    ) {
      await scope
        .locator("text=/^Change Folder$/i")
        .first()
        .click({ timeout: 60000 })
        .catch(() => {});
    } else {
      const menuOpened = await openDocumentListOverflowMenu(scope);
      if (menuOpened) {
        const clicked = await tryClickAnyAction(scope, [
          /^Change Folder$/i,
          /^Move to Folder$/i,
          /^Move Folder$/i,
        ]);
        if (!clicked && attempt === 2) {
          throw new Error('Could not find "Change Folder" action in document menu.');
        }
      } else if (attempt === 2) {
        throw new Error('Could not open document list menu for "Change Folder".');
      }
    }

    const dialogVisible = await Promise.any([
      scope
        .locator("text=/^Change Document Folder$/i")
        .first()
        .waitFor({ timeout: 4000 })
        .then(() => true),
      scope
        .locator("#folderselection")
        .first()
        .waitFor({ timeout: 4000 })
        .then(() => true),
      scope
        .locator("input#change_folder_confirm")
        .first()
        .waitFor({ timeout: 4000 })
        .then(() => true),
    ]).catch(() => false);

    if (dialogVisible) return;
    await dismissTransientBlockingModals(scope, "change folder dialog did not appear");
  }

  await saveMoveDebugArtifacts(scope, "change-folder-dialog-not-visible");
  throw new Error(
    "Change Folder dialog did not appear. " +
      "Debug files: clean-move-debug-change-folder-dialog-not-visible.json/png"
  );
}

async function moveToFolder(scope, folderName) {
  scope = await resolveActionScope(scope);
  await dismissTransientBlockingModals(scope, "before destination selection");

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

  const hasDataTarget = (await targetByDataName.count().catch(() => 0)) > 0;
  const hasFallbackTarget = (await targetFallback.count().catch(() => 0)) > 0;

  if (!hasDataTarget && !hasFallbackTarget) {
    await saveMoveDebugArtifacts(scope, "destination-folder-not-visible");
    throw new Error(
      `Destination folder "${folderName}" not visible in change-folder dialog. ` +
        "Debug files: clean-move-debug-destination-folder-not-visible.json/png"
    );
  }

  if (hasDataTarget) {
    await targetByDataName.click({ timeout: 60000 });
  } else {
    await targetFallback.waitFor({ timeout: 60000 });
    await targetFallback.click({ timeout: 60000 });
  }

  const confirmById = scope.locator("input#change_folder_confirm").first();
  if ((await confirmById.count().catch(() => 0)) > 0) {
    await confirmById.click({ timeout: 60000 });
  } else {
    await scope
      .locator('button:has-text("Confirm"), button:has-text("Move"), input[value="Confirm"]')
      .first()
      .click({ timeout: 60000 });
  }

  await dismissTransientBlockingModals(scope, "after destination confirm");
}

async function openDocumentListOverflowMenu(scope) {
  const candidates = [
    '#document_list_header button:has-text("...")',
    '#document_list_header [role="button"]:has-text("...")',
    '#document_list_header button:has-text("…")',
    '#document_list_header [role="button"]:has-text("…")',
    '#document_list_header button:has-text("⋯")',
    '#document_list_header [role="button"]:has-text("⋯")',
    '#document_list button:has-text("...")',
    '#document_list [role="button"]:has-text("...")',
    '#document_list button:has-text("…")',
    '#document_list [role="button"]:has-text("…")',
    '#document_list button:has-text("⋯")',
    '#document_list [role="button"]:has-text("⋯")',
    'button:has-text("...")',
    '[role="button"]:has-text("...")',
    'button:has-text("…")',
    '[role="button"]:has-text("…")',
    'button:has-text("⋯")',
    '[role="button"]:has-text("⋯")',
    '[aria-label*="more" i]',
    '[aria-label*="menu" i]',
    '[title*="more" i]',
    '[title*="menu" i]',
    '[class*="ellipsis" i]',
    '[class*="kebab" i]',
    '[class*="more" i]',
  ];

  for (const selector of candidates) {
    const loc = scope.locator(selector);
    const count = Math.min(await loc.count().catch(() => 0), 8);
    for (let i = 0; i < count; i++) {
      const item = loc.nth(i);
      const visible = await item.isVisible({ timeout: 200 }).catch(() => false);
      if (!visible) continue;
      await item.click().catch(() => {});
      await waitForTimeout(scope, 180);
      return true;
    }
  }

  return false;
}

async function tryClickAnyAction(scope, regexList) {
  for (const regex of regexList) {
    const candidate = scope.getByText(regex).first();
    const visible = await candidate.isVisible({ timeout: 300 }).catch(() => false);
    if (!visible) continue;
    try {
      await candidate.click({ timeout: 5000 });
      return true;
    } catch (_) {
      continue;
    }
  }
  return false;
}

async function tryEnableSelectModeFromMenu(scope) {
  const checkboxCandidates = [
    'label:has-text("Select Mode") input[type="checkbox"]',
    'li:has-text("Select Mode") input[type="checkbox"]',
    'div:has-text("Select Mode") input[type="checkbox"]',
    '[role="menuitemcheckbox"]:has-text("Select Mode") input[type="checkbox"]',
  ];

  for (const selector of checkboxCandidates) {
    const cb = scope.locator(selector).first();
    const exists = (await cb.count().catch(() => 0)) > 0;
    if (!exists) continue;

    const checked = await cb.isChecked().catch(() => false);
    if (!checked) {
      await cb.click({ timeout: 5000, force: true }).catch(() => {});
      await waitForTimeout(scope, 120);

      const checkedAfterClick = await cb.isChecked().catch(() => false);
      if (!checkedAfterClick) {
        await cb.press("Space").catch(() => {});
      }
    }
    if (await isSelectModeEnabled(scope)) return true;
  }

  // Some tenants require clicking the menu row left-edge where the checkbox is rendered.
  const rowClick = await clickSelectModeRowCheckboxArea(scope);
  if (rowClick) {
    await waitForTimeout(scope, 150);
    if (await isSelectModeEnabled(scope)) return true;
    return true;
  }

  const clicked = await tryClickAnyAction(scope, [
    /^Select Mode$/i,
    /^Select Documents?$/i,
    /^Multi[- ]?select$/i,
    /^Select$/i,
  ]);
  if (!clicked) return false;

  await waitForTimeout(scope, 120);
  if (await isSelectModeEnabled(scope)) return true;

  // Last fallback: if we successfully clicked a Select Mode action but this tenant does
  // not expose checked-state/checkboxes in DOM, continue optimistically.
  return true;
}

async function isSelectModeEnabled(scope) {
  const rowCheckboxCount = await scope
    .locator('#document_list input[type="checkbox"]')
    .count()
    .catch(() => 0);
  if (rowCheckboxCount > 0) return true;

  const menuCheckboxCandidates = [
    'label:has-text("Select Mode") input[type="checkbox"]',
    'li:has-text("Select Mode") input[type="checkbox"]',
    'div:has-text("Select Mode") input[type="checkbox"]',
    '[role="menuitemcheckbox"]:has-text("Select Mode") input[type="checkbox"]',
  ];

  for (const selector of menuCheckboxCandidates) {
    const cb = scope.locator(selector).first();
    const count = await cb.count().catch(() => 0);
    if (!count) continue;
    const checked = await cb.isChecked().catch(() => false);
    if (checked) return true;
  }

  const menuAriaChecked = await scope
    .locator(
      [
        '[role="menuitemcheckbox"][aria-checked="true"]:has-text("Select Mode")',
        '[aria-checked="true"]:has-text("Select Mode")',
      ].join(", ")
    )
    .count()
    .catch(() => 0);
  if (menuAriaChecked > 0) return true;

  const activeSelectMode = await scope
    .locator(
      [
        '[class*="select" i][class*="active" i]:has-text("Select Mode")',
        '[class*="active" i]:has-text("Select Mode")',
      ].join(", ")
    )
    .count()
    .catch(() => 0);

  return activeSelectMode > 0;
}

async function dismissTransientBlockingModals(scope, reason = "unknown") {
  const page = getScopePage(scope);
  const deadline = Date.now() + 2500;
  let dismissedAny = false;

  while (Date.now() < deadline) {
    const ajaxTitle = page.locator("text=/AJAX Issue/i").first();
    const ajaxVisible = await ajaxTitle.isVisible({ timeout: 120 }).catch(() => false);

    const genericModal = page
      .locator(
        [
          '.modal:visible',
          '.bootbox:visible',
          '.alertify.ajs-in',
          '[role="dialog"]:visible',
        ].join(", ")
      )
      .first();
    const genericVisible = await genericModal.isVisible({ timeout: 120 }).catch(() => false);

    if (!ajaxVisible && !genericVisible) break;

    const okBtn = page
      .locator(
        [
          'button:has-text("Ok")',
          'button:has-text("OK")',
          'button:has-text("Close")',
          'button:has-text("Confirm")',
          'input[value="OK"]',
          'input[value="Ok"]',
        ].join(", ")
      )
      .first();

    if ((await okBtn.count().catch(() => 0)) > 0) {
      await okBtn.click({ force: true }).catch(() => {});
      dismissedAny = true;
    } else {
      await page.keyboard.press("Escape").catch(() => {});
      dismissedAny = true;
    }

    await page.waitForTimeout(150);
  }

  if (dismissedAny) {
    console.log(`⚠ Dismissed blocking modal (${reason})`);
  }
}

async function clickSelectModeRowCheckboxArea(scope) {
  const page = getScopePage(scope);
  const rowSelectors = [
    '#document_list_header label:has-text("Select Mode")',
    '#document_list_header [role="menuitemcheckbox"]:has-text("Select Mode")',
    '#document_list_header li:has-text("Select Mode")',
    '#document_list_header div:has-text("Select Mode")',
    '#document_list label:has-text("Select Mode")',
    '#document_list [role="menuitemcheckbox"]:has-text("Select Mode")',
    '#document_list li:has-text("Select Mode")',
    '#document_list div:has-text("Select Mode")',
    'label:has-text("Select Mode")',
    '[role="menuitemcheckbox"]:has-text("Select Mode")',
    'li:has-text("Select Mode")',
    'div:has-text("Select Mode")',
  ];

  for (const selector of rowSelectors) {
    const rows = scope.locator(selector);
    const count = Math.min(await rows.count().catch(() => 0), 8);
    for (let i = 0; i < count; i++) {
      const row = rows.nth(i);
      const visible = await row.isVisible({ timeout: 200 }).catch(() => false);
      if (!visible) continue;

      const box = await row.boundingBox().catch(() => null);
      if (box) {
        const x = box.x + Math.min(14, Math.max(6, box.width * 0.12));
        const y = box.y + Math.max(6, box.height / 2);
        await page.mouse.click(x, y).catch(() => {});
        await waitForTimeout(scope, 80);
      }

      await row.click({ timeout: 3000, force: true }).catch(() => {});
      return true;
    }
  }

  return false;
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

async function saveMoveDebugArtifacts(scope, reason) {
  try {
    const page = getScopePage(scope);
    const tag = sanitizeFileName(reason || "move");
    const jsonFile = `clean-move-debug-${tag}.json`;
    const screenshotFile = `clean-move-debug-${tag}.png`;

    const info = {
      createdAt: new Date().toISOString(),
      reason,
      pageUrl: page.url(),
      scopeType: scope === page ? "page" : "frame",
      scopeUrl: typeof scope.url === "function" ? scope.url() : page.url(),
      selectors: {
        documentList: await scope.locator("#document_list").count().catch(() => 0),
        checkboxes: await scope
          .locator('#document_list input[type="checkbox"]')
          .count()
          .catch(() => 0),
        actionSelectMode: await scope
          .locator("a#action_selectmode, button#action_selectmode")
          .count()
          .catch(() => 0),
        actionChangeFolder: await scope
          .locator("a#action_changefolder, button#action_changefolder")
          .count()
          .catch(() => 0),
      },
      visibleMenuLabels: await scope
        .locator("#document_list button, #document_list a")
        .evaluateAll((nodes) => {
          const labels = [];
          for (const n of nodes) {
            const text = (n.textContent || "").replace(/\s+/g, " ").trim();
            if (!text) continue;
            if (!labels.includes(text)) labels.push(text);
            if (labels.length >= 100) break;
          }
          return labels;
        })
        .catch(() => []),
      visibleActionTexts: await scope
        .locator("label, a, button, li, div")
        .evaluateAll((nodes) => {
          const labels = [];
          for (const n of nodes) {
            const style = window.getComputedStyle(n);
            if (!style || style.display === "none" || style.visibility === "hidden") {
              continue;
            }
            const text = (n.textContent || "").replace(/\s+/g, " ").trim();
            if (!text) continue;
            if (text.length > 60) continue;
            if (!labels.includes(text)) labels.push(text);
            if (labels.length >= 200) break;
          }
          return labels;
        })
        .catch(() => []),
    };

    fs.writeFileSync(path.join(process.cwd(), jsonFile), JSON.stringify(info, null, 2), "utf8");
    await page.screenshot({ path: screenshotFile, fullPage: true }).catch(() => {});
  } catch (_) {}
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

function promptText(q, defaultValue = "") {
  return inquirer
    .prompt([
      {
        type: "input",
        name: "value",
        message: q,
        default: defaultValue,
      },
    ])
    .then((ans) => (ans.value || "").trim());
}

module.exports = cleanBetterLetterProcessing;
