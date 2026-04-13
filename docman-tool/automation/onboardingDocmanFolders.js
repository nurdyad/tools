const DEFAULT_FOLDER_NAMES = [
  "zz BL Filing. Do not touch",
  "BetterLetter: Rejected",
  "zz BL Processing. Do not touch",
  "zz BL Input. Do not touch",
];

const DOCMAN_HOST_MATCH = "docman.thirdparty.nhs.uk";
const DOCUMENT_FOLDERS_PATH = "/Admin/DocumentFolders";

const SETTINGS_MENU_SELECTORS = [
  'a:has-text("Settings")',
  'button:has-text("Settings")',
  '[aria-label*="settings" i]',
  '[title*="settings" i]',
  'a[href*="/Settings" i]',
  'button[title*="Settings" i]',
];

const SETTINGS_CONSOLE_SELECTORS = [
  'text=/Settings Console/i',
  'text=/Document Folders/i',
  'text=/Pre-defined Filing Codes/i',
];

const FILING_MENU_SELECTORS = [
  'a:has-text("Filing")',
  'button:has-text("Filing")',
  'li:has-text("Filing")',
  'text=/^\\s*Filing\\s*$/i',
];

const DOCUMENT_FOLDERS_MENU_SELECTORS = [
  'a:has-text("Document Folders")',
  'button:has-text("Document Folders")',
  'li:has-text("Document Folders") a',
  'li:has-text("Document Folders") button',
  'li:has-text("Document Folders")',
];

const DOCUMENT_FOLDERS_SCREEN_SELECTORS = [
  'text=/Document Folders List/i',
  'text=/Edit Folders\\s*Filing/i',
  'text=/Top Level Folder/i',
  'text=/Back to List/i',
  'text=/New Folder Name/i',
];

const FILING_SECTION_ROW_SELECTORS = [
  'xpath=//table//tr[./td[normalize-space()="Filing"]]',
  'xpath=//table//tr[.//*[normalize-space()="Filing"]]',
  'xpath=//main//table//*[self::a or self::button][normalize-space()="Filing"]',
];

const FILING_SECTION_ACTION_SELECTORS = [
  'xpath=//table//tr[.//*[normalize-space()="Filing"]]//*[self::a or self::button]',
  'xpath=//main//*[contains(@class,"selected") and .//*[normalize-space()="Filing"]]//*[self::a or self::button]',
];

const EDIT_FOLDERS_SELECTORS = [
  'text=/Edit Folders\\s*Filing/i',
  'text=/Top Level Folder/i',
  'text=/Selected Folder/i',
];

const TOP_LEVEL_FOLDER_SELECTORS = [
  'text=/^\\s*Top Level Folder\\s*$/i',
  'a:has-text("Top Level Folder")',
  'li:has-text("Top Level Folder")',
  'div:has-text("Top Level Folder")',
];

const ADD_BUTTON_SELECTORS = [
  'a#addFolder',
  '#addFolder',
  'xpath=//*[contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"selected folder")]//a[normalize-space()="Add"]',
  'xpath=//*[contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"selected folder")]//button[normalize-space()="Add"]',
  'xpath=//*[contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"selected folder")]/following::*[(self::a or self::button) and normalize-space()="Add"][1]',
  'a:has-text("Add")',
  'button:has-text("Add")',
  'input[value="Add"]',
];

const NEW_FOLDER_INPUT_SELECTORS = [
  'xpath=//label[contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"new folder name")]/following::input[1]',
  'input[placeholder*="new folder" i]',
  'input[name*="folder" i]',
  'input[id*="folder" i]',
];

const CONFIRM_BUTTON_SELECTORS = [
  '#manageFolder #addFolderConfirm:not(.disabled):not([disabled])',
  'a#addFolderConfirm:not(.disabled):not([disabled])',
  '#addFolderConfirm:not(.disabled):not([disabled])',
  'a:has-text("Confirm"):not(.disabled):not([disabled])',
  'button:has-text("Confirm"):not(.disabled):not([disabled])',
  'input[value="Confirm"]:not([disabled])',
  'a#addFolderConfirm',
  '#addFolderConfirm',
  'a:has-text("Confirm")',
  'button:has-text("Confirm")',
  'input[value="Confirm"]',
];

const DOCMAN_LOGIN_PAGE_SELECTORS = [
  'input[name="Username"]',
  'input[name="Password"]',
  'button:has-text("Login")',
  'button:has-text("Sign in")',
  'text=/forgot password/i',
];

async function onboardingDocmanFolders({ page, inputFolderName, timeouts = {}, logger = null } = {}) {
  if (!page) {
    throw new Error("onboardingDocmanFolders requires a Playwright page.");
  }

  const timeoutMs = Number(timeouts?.selectorMs) || 60000;
  const folderNames = resolveOnboardingFolders(inputFolderName);

  console.log("➡ Opening Docman Settings → Filing → Document Folders");
  console.log(`📁 Folders to ensure: ${folderNames.join(" | ")}`);

  const stepToken = logger?.startStep?.("onboarding_docman_folders_internal", {
    folderCount: folderNames.length,
  });
  const summary = {
    requested: folderNames.length,
    created: 0,
    existing: 0,
    skipped: 0,
    failed: 0,
  };

  try {
    const result = await runWithModalWatcher(page, async () => {
      await dismissBlockingModals(page, "before onboarding start", { quiet: true });
      await openDocumentFoldersEditor(page, timeoutMs);
      const folderCache = { names: new Set(), lastSyncAt: 0 };
      await refreshFolderCache(page, folderCache, true);

      for (const folderName of folderNames) {
        const name = String(folderName || "").trim();
        if (!name) {
          summary.skipped += 1;
          continue;
        }

        try {
          await ensureOnEditFoldersScreen(page, timeoutMs);
          const status = await ensureFolder(page, name, timeoutMs, folderCache);
          if (status === "created") summary.created += 1;
          else if (status === "exists") summary.existing += 1;
          else summary.skipped += 1;
        } catch (error) {
          summary.failed += 1;
          throw new Error(`Folder "${name}" failed: ${error?.message || String(error)}`);
        }
      }

      await dismissBlockingModals(page, "after onboarding finish", { quiet: true });
      return { ...summary, folderNames };
    });

    printOnboardingSummary(result);
    console.log(
      `✔ Onboarding folders complete (created=${result.created}, already-existed=${result.existing}, skipped=${result.skipped}, failed=${result.failed})`
    );
    logger?.endStep?.(stepToken, {
      status: "ok",
      requested: result.requested,
      created: result.created,
      existing: result.existing,
      skipped: result.skipped,
      failed: result.failed,
    });
    return result;
  } catch (error) {
    printOnboardingSummary(summary);
    logger?.endStep?.(stepToken, {
      status: "error",
      requested: summary.requested,
      created: summary.created,
      existing: summary.existing,
      skipped: summary.skipped,
      failed: summary.failed,
      errorMessage: error?.message || String(error),
    });
    const debug = await captureOnboardingDebug(page, "failed");
    const msg = error?.message || String(error);
    throw new Error(`${msg}. Saved debug files: ${debug.screenshotPath}, ${debug.htmlPath}`);
  }
}

function resolveOnboardingFolders(inputFolderName) {
  const folder4 = String(inputFolderName || DEFAULT_FOLDER_NAMES[3]).trim() || DEFAULT_FOLDER_NAMES[3];
  return [...DEFAULT_FOLDER_NAMES.slice(0, 3), folder4];
}

function printOnboardingSummary(summary = {}) {
  const requested = Number(summary.requested) || 0;
  const created = Number(summary.created) || 0;
  const existing = Number(summary.existing) || 0;
  const skipped = Number(summary.skipped) || 0;
  const failed = Number(summary.failed) || 0;

  console.log("📊 ONBOARDING SUMMARY");
  console.log(`  Requested: ${requested}`);
  console.log(`  Created: ${created}`);
  console.log(`  Already existed: ${existing}`);
  console.log(`  Skipped: ${skipped}`);
  console.log(`  Failed: ${failed}`);
}

async function openDocumentFoldersEditor(page, timeoutMs) {
  await assertDocmanSessionReady(page, "opening Document Folders");
  if (await isAnyVisible(page, EDIT_FOLDERS_SELECTORS, 1200)) {
    return;
  }

  await tryOpenDocumentFoldersByUrl(page, timeoutMs);
  await assertDocmanSessionReady(page, "opening Document Folders");
  if (await isAnyVisible(page, EDIT_FOLDERS_SELECTORS, 1200)) {
    return;
  }

  if (!(await isAnyVisible(page, SETTINGS_CONSOLE_SELECTORS, 1200))) {
    await clickAny(page, SETTINGS_MENU_SELECTORS, {
      timeoutMs: Math.min(timeoutMs, 10000),
      description: "Settings launcher",
      throwOnMissing: false,
    });
  }

  await waitAny(page, SETTINGS_CONSOLE_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 20000),
    description: "Settings Console",
  });

  if (!(await isAnyVisible(page, DOCUMENT_FOLDERS_SCREEN_SELECTORS, 1200))) {
    console.log("↳ Opening Filing menu in Settings Console");
    await clickAny(page, FILING_MENU_SELECTORS, {
      timeoutMs: Math.min(timeoutMs, 20000),
      description: "Filing menu",
      throwOnMissing: false,
    });

    console.log("↳ Opening Document Folders");
    await clickDocumentFoldersMenu(page, {
      timeoutMs,
      description: "Document Folders menu",
    });
  }

  await waitAny(page, DOCUMENT_FOLDERS_SCREEN_SELECTORS, {
    timeoutMs,
    description: "Document Folders screen",
  });

  if (!(await isAnyVisible(page, EDIT_FOLDERS_SELECTORS, 1200))) {
    console.log("↳ Opening Filing section row in Document Folders list");
    await openFilingSectionEditor(page, timeoutMs);
  }

  await waitAny(page, EDIT_FOLDERS_SELECTORS, {
    timeoutMs,
    description: "Edit Folders (Filing) screen",
  });
}

async function clickDocumentFoldersMenu(page, options = {}) {
  const { timeoutMs = 15000, description = "Document Folders menu" } = options;
  const found = await waitAny(page, DOCUMENT_FOLDERS_MENU_SELECTORS, {
    timeoutMs,
    description,
  });

  const href = String((await found.locator.getAttribute("href").catch(() => "")) || "").trim();
  await found.locator.scrollIntoViewIfNeeded().catch(() => {});
  await found.locator.click({ timeout: 5000 }).catch(async () => {
    const childAction = found.locator.locator("a, button").first();
    if ((await childAction.count().catch(() => 0)) > 0) {
      await childAction.click({ timeout: 5000 }).catch(() => {});
      return;
    }
    await found.locator.click({ timeout: 5000, force: true }).catch(() => {});
  });
  await sleep(200);

  if (await isAnyVisible(page, DOCUMENT_FOLDERS_SCREEN_SELECTORS, 2500)) return;
  if (!href) return;

  const scopeBase = getScopeUrl(found.scope, page);
  const absolute = href.startsWith("http")
    ? href
    : new URL(href, scopeBase).toString();
  await page.goto(absolute, {
    waitUntil: "domcontentloaded",
    timeout: timeoutMs,
  }).catch(() => {});
}

async function openFilingSectionEditor(page, timeoutMs) {
  const filing = await waitAny(page, FILING_SECTION_ROW_SELECTORS, {
    timeoutMs,
    description: "Filing section row",
  });

  const attempts = [
    async () => {
      await filing.locator.click({ timeout: 5000 }).catch(() => {});
    },
    async () => {
      await filing.locator.dblclick({ timeout: 5000 }).catch(() => {});
    },
    async () => {
      const action = await findAnyVisible(page, FILING_SECTION_ACTION_SELECTORS, 1200);
      if (action) {
        await action.locator.click({ timeout: 5000 }).catch(() => {});
      }
    },
    async () => {
      await filing.locator.focus().catch(() => {});
      await getScopePage(filing.scope).keyboard.press("Enter").catch(() => {});
    },
  ];

  for (const action of attempts) {
    await action();
    await sleep(250);
    await dismissBlockingModals(page, "after filing row action", { quiet: true });
    if (await isAnyVisible(page, EDIT_FOLDERS_SELECTORS, 2500)) return;
  }

  throw new Error("Filing section row clicked but Edit Folders screen did not open");
}

async function ensureFolder(page, folderName, timeoutMs, folderCache) {
  const name = String(folderName || "").trim();
  if (!name) return "exists";

  if (await hasFolderVisible(page, name, folderCache)) {
    console.log(`↳ Folder exists: ${name}`);
    return "exists";
  }

  await dismissBlockingModals(page, `before creating "${name}"`, { quiet: true });

  await clickAny(page, TOP_LEVEL_FOLDER_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 15000),
    description: "Top Level Folder",
  });

  await clickAny(page, ADD_BUTTON_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 10000),
    description: "Add folder button",
  });

  const input = await waitAny(page, NEW_FOLDER_INPUT_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 10000),
    description: "New Folder Name input",
  });

  await input.locator.click({ clickCount: 3 }).catch(() => {});
  await input.locator.fill("").catch(() => {});
  await input.locator.fill(name).catch(async () => {
    await input.locator.type(name).catch(() => {});
  });

  const typed = String((await input.locator.inputValue().catch(() => "")) || "").trim();
  if (typed !== name) {
    await input.locator.fill(name).catch(() => {});
  }

  await input.locator.evaluate((el) => {
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }).catch(() => {});
  await input.locator.focus().catch(() => {});
  await input.locator.press("End").catch(() => {});
  await input.locator.type(" ", { delay: 10 }).catch(() => {});
  await input.locator.press("Backspace").catch(() => {});

  try {
    await clickAny(page, CONFIRM_BUTTON_SELECTORS, {
      timeoutMs: Math.min(timeoutMs, 10000),
      description: `Confirm create folder "${name}"`,
    });
  } catch (_) {
    await input.locator.press("Enter").catch(() => {});
    await page.evaluate(() => {
      const btn = document.querySelector("#addFolderConfirm");
      if (btn && typeof btn.click === "function") btn.click();
    }).catch(() => {});
  }

  const status = await waitForFolderResult(page, name, timeoutMs, folderCache);
  console.log(`↳ Folder ${status}: ${name}`);
  return status;
}

async function waitForFolderResult(page, folderName, timeoutMs, folderCache) {
  const startedAt = Date.now();
  const settleMs = Math.min(timeoutMs, 8000);

  while (Date.now() - startedAt < settleMs) {
    if (await hasFolderVisible(page, folderName, folderCache)) return "created";

    const pageText = await readAllScopeText(page);
    if (/already exists|duplicate|cannot create|exists/i.test(pageText)) {
      await dismissBlockingModals(page, "folder already exists", { quiet: true });
      await refreshFolderCache(page, folderCache, true);
      return "exists";
    }

    await dismissBlockingModals(page, "waiting for folder result", { quiet: true });
    await sleep(150);
  }

  throw new Error(`Timed out waiting for folder creation result: ${folderName}`);
}

async function hasFolderVisible(page, folderName, folderCache) {
  const normalized = normalizeFolderName(folderName);
  if (!normalized) return false;
  await refreshFolderCache(page, folderCache);
  return folderCache?.names?.has(normalized) || false;
}

async function runWithModalWatcher(page, task) {
  let stop = false;
  const watcher = (async () => {
    while (!stop) {
      await dismissBlockingModals(page, "background", { quiet: true }).catch(() => {});
      await sleep(300);
    }
  })();

  try {
    return await task();
  } finally {
    stop = true;
    await watcher.catch(() => {});
    await dismissBlockingModals(page, "final", { quiet: true }).catch(() => {});
  }
}

async function dismissBlockingModals(page, reason = "unknown", options = {}) {
  const quiet = Boolean(options.quiet);
  let dismissed = 0;

  const modalRootSelectors = [
    ".alertify.ajs-in",
    ".alertify.ajs-fade.ajs-in",
    ".swal2-popup",
    '[role="dialog"]',
  ];
  const closeSelectors = [
    'button:has-text("OK")',
    'button:has-text("Ok")',
    'button:has-text("No")',
    'button:has-text("Cancel")',
    'button:has-text("Close")',
    'button:has-text("Dismiss")',
    'a:has-text("OK")',
    'a:has-text("No")',
    'a:has-text("Cancel")',
    'a:has-text("Close")',
    'button[aria-label="Close"]',
    'button:has-text("×")',
  ];

  for (const scope of getScopes(page)) {
    for (const modalSelector of modalRootSelectors) {
      const modals = scope.locator(modalSelector);
      const count = Math.min(await modals.count().catch(() => 0), 4);

      for (let i = 0; i < count; i += 1) {
        const modal = modals.nth(i);
        const visible = await modal.isVisible({ timeout: 100 }).catch(() => false);
        if (!visible) continue;

        let clicked = false;
        for (const closeSelector of closeSelectors) {
          const btn = modal.locator(closeSelector).first();
          if ((await btn.count().catch(() => 0)) === 0) continue;
          await btn.click({ force: true, timeout: 1500 }).catch(() => {});
          clicked = true;
          dismissed += 1;
          break;
        }

        if (!clicked) {
          await getScopePage(scope).keyboard.press("Escape").catch(() => {});
          dismissed += 1;
        }
      }
    }
  }

  if (dismissed > 0 && !quiet) {
    console.log(`⚠ Dismissed ${dismissed} blocking dialog(s) (${reason})`);
  }

  return dismissed;
}

async function assertDocmanSessionReady(page, reason = "onboarding") {
  if (await isLikelyDocmanLoginPage(page)) {
    throw new Error(`Docman appears to be on the login page while ${reason}`);
  }

  const currentUrl = String(page.url() || "");
  if (!currentUrl) return;

  let host = "";
  try {
    host = new URL(currentUrl).host.toLowerCase();
  } catch (_) {
    host = "";
  }

  if (host && !host.includes(DOCMAN_HOST_MATCH)) {
    throw new Error(`Unexpected page while ${reason}: ${currentUrl}`);
  }
}

async function ensureOnEditFoldersScreen(page, timeoutMs) {
  await assertDocmanSessionReady(page, "running onboarding");
  if (await isAnyVisible(page, EDIT_FOLDERS_SELECTORS, 1200)) return;
  await openDocumentFoldersEditor(page, timeoutMs);
}

async function isLikelyDocmanLoginPage(page) {
  const url = String(page.url() || "");
  if (/\/account\/login/i.test(url)) return true;
  return isAnyVisible(page, DOCMAN_LOGIN_PAGE_SELECTORS, 500);
}

async function tryOpenDocumentFoldersByUrl(page, timeoutMs) {
  const currentUrl = String(page.url() || "");
  let targetUrl = "";
  try {
    const parsed = new URL(currentUrl);
    if (!String(parsed.host || "").toLowerCase().includes(DOCMAN_HOST_MATCH)) return false;
    targetUrl = `${parsed.origin}${DOCUMENT_FOLDERS_PATH}`;
  } catch (_) {
    return false;
  }

  if (!targetUrl) return false;
  if (currentUrl.startsWith(targetUrl)) return true;

  await page
    .goto(targetUrl, {
      waitUntil: "domcontentloaded",
      timeout: Math.min(timeoutMs, 20000),
    })
    .catch(() => {});

  return true;
}

async function refreshFolderCache(page, folderCache, force = false) {
  if (!folderCache) return new Set();
  const now = Date.now();
  if (!force && folderCache.lastSyncAt && now - folderCache.lastSyncAt < 250) {
    return folderCache.names;
  }

  const names = new Set();
  for (const scope of getScopes(page)) {
    const scopeNames = await scope
      .evaluate(() => {
        const out = [];

        const treeLinks = document.querySelectorAll("#folderTree a[data-foldername]");
        for (const link of treeLinks) {
          const raw = String(link.getAttribute("data-foldername") || "").trim();
          if (raw) out.push(raw);
        }

        const flatInput = document.querySelector("#foldersList");
        const flatRaw = String(flatInput?.value || "").trim();
        if (flatRaw) {
          try {
            const parsed = JSON.parse(flatRaw);
            if (Array.isArray(parsed)) {
              for (const row of parsed) {
                const raw = String(row?.FolderName || "").trim();
                if (raw) out.push(raw);
              }
            }
          } catch (_) {}
        }

        return out;
      })
      .catch(() => []);

    for (const name of scopeNames) {
      const normalized = normalizeFolderName(name);
      if (normalized) names.add(normalized);
    }
  }

  folderCache.names = names;
  folderCache.lastSyncAt = now;
  return names;
}

function normalizeFolderName(name) {
  return String(name || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

async function clickAny(page, selectors, options = {}) {
  const { timeoutMs = 15000, description = "element", throwOnMissing = true } = options;
  const found = await findAnyVisible(page, selectors, timeoutMs);
  if (!found) {
    if (!throwOnMissing) return null;
    throw new Error(`${description} not found (selectors=${selectors.join(" | ")})`);
  }

  await found.locator.scrollIntoViewIfNeeded().catch(() => {});
  await found.locator.click({ timeout: 5000 }).catch(async () => {
    await found.locator.click({ timeout: 5000, force: true }).catch(() => {});
  });
  await sleep(150);
  return found;
}

async function waitAny(page, selectors, options = {}) {
  const { timeoutMs = 15000, description = "element" } = options;
  const found = await findAnyVisible(page, selectors, timeoutMs);
  if (!found) {
    throw new Error(`${description} not visible (selectors=${selectors.join(" | ")})`);
  }
  return found;
}

async function isAnyVisible(page, selectors, timeoutMs = 700) {
  return Boolean(await findAnyVisible(page, selectors, timeoutMs));
}

async function findAnyVisible(page, selectors, timeoutMs = 15000) {
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    for (const scope of getScopes(page)) {
      for (const selector of selectors) {
        const candidates = scope.locator(selector);
        const count = await candidates.count().catch(() => 0);
        if (!count) continue;

        const inspectCount = Math.min(count, 8);
        for (let i = 0; i < inspectCount; i += 1) {
          const locator = candidates.nth(i);
          const visible = await locator.isVisible({ timeout: 150 }).catch(() => false);
          if (!visible) continue;
          return { scope, locator, selector };
        }
      }
    }
    await sleep(200);
  }

  return null;
}

function getScopes(page) {
  const scopes = [page];
  for (const frame of page.frames()) {
    if (frame === page.mainFrame()) continue;
    scopes.push(frame);
  }
  return scopes;
}

function getScopePage(scope) {
  return typeof scope.page === "function" ? scope.page() : scope;
}

function getScopeUrl(scope, page) {
  try {
    if (scope && typeof scope.url === "function") {
      const u = String(scope.url() || "").trim();
      if (u) return u;
    }
  } catch (_) {}
  return page.url();
}

async function readAllScopeText(page) {
  const chunks = [];
  for (const scope of getScopes(page)) {
    const text = await scope.locator("body").innerText().catch(() => "");
    if (text) chunks.push(text);
  }
  return chunks.join("\n");
}

function escapeRegExp(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function captureOnboardingDebug(page, suffix = "debug") {
  const safe = String(suffix || "debug").replace(/[^a-z0-9_-]+/gi, "-");
  const screenshotPath = `onboarding-${safe}.png`;
  const htmlPath = `onboarding-${safe}.html`;

  await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
  const html = await page.content().catch(() => "");
  if (html) {
    const fs = require("fs");
    fs.writeFileSync(htmlPath, html, "utf8");
  }

  return { screenshotPath, htmlPath };
}

module.exports = onboardingDocmanFolders;
module.exports.resolveOnboardingFolders = resolveOnboardingFolders;
module.exports.DEFAULT_FOLDER_NAMES = DEFAULT_FOLDER_NAMES;
