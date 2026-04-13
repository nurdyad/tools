const fs = require("fs");

const DOCMAN_HOST_MATCH = "docman.thirdparty.nhs.uk";
const USER_GROUPS_LIST_PATH = "/Admin/RecipientGroups/RecipientGroupList";

const DOCMAN_LOGIN_PAGE_SELECTORS = [
  'input[name="Username"]',
  'input[name="Password"]',
  'button:has-text("Login")',
  'button:has-text("Sign in")',
  'text=/forgot password/i',
  'text=/Sign in to Continue/i',
];

const USER_GROUPS_MENU_SELECTORS = [
  'a[href*="/Admin/RecipientGroups"]',
  'a:has-text("User Groups")',
  'button:has-text("User Groups")',
  'li:has-text("User Groups") a',
];

const USER_GROUP_LIST_SCREEN_SELECTORS = [
  'text=/User Groups List/i',
  'text=/Description/i',
  'a:has-text("Create")',
  'button:has-text("Create")',
  'input[value="Create"]',
];

const LIST_CREATE_BUTTON_SELECTORS = [
  'a:has-text("Create")',
  'button:has-text("Create")',
  'input[value="Create"]',
];

const CREATE_GROUP_SCREEN_SELECTORS = [
  'text=/Create New/i',
  'text=/Group Details/i',
  'text=/Group Members/i',
  'button:has-text("Manage Members")',
  'a:has-text("Manage Members")',
  'input[value="Manage Members"]',
];

const GROUP_DESCRIPTION_INPUT_SELECTORS = [
  'xpath=//label[contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"description")]/following::input[1]',
  'input[name*="Description" i]',
  'input[id*="Description" i]',
  'input[placeholder*="description" i]',
  'input[type="text"]',
];

const MANAGE_MEMBERS_BUTTON_SELECTORS = [
  'button:has-text("Manage Members")',
  'a:has-text("Manage Members")',
  'input[value="Manage Members"]',
];

const USER_SEARCH_TAB_SELECTORS = [
  '#usersearch_modal a[href="#tab2"]',
  '#usersearch_modal li:has(a[href="#tab2"])',
  'button:has-text("User Search")',
  'a:has-text("User Search")',
  '[role="tab"]:has-text("User Search")',
  'li:has-text("User Search")',
];

const MODAL_SEARCH_INPUT_SELECTORS = [
  '#usersearch_modal #SearchTerm',
  'input[placeholder*="search" i]',
  'input[aria-label*="search" i]',
  'input[name*="search" i]',
  'input[id*="search" i]',
  'input[type="search"]',
  'xpath=.//label[contains(translate(normalize-space(.),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"search")]/following::input[1]',
  'input:not([type="hidden"]):not([type="checkbox"]):not([type="radio"])',
];

const MODAL_CONFIRM_BUTTON_SELECTORS = [
  '#usersearch_confirm:not(.disabled):not([disabled])',
  '#usersearch_confirm',
  'button:has-text("Confirm")',
  'a:has-text("Confirm")',
  'input[value="Confirm"]',
];

const PAGE_CREATE_BUTTON_SELECTORS = [
  'button:has-text("Create")',
  'a:has-text("Create")',
  'input[value="Create"]',
];

async function createDocmanUserGroup({
  page,
  groupName,
  usernames,
  timeouts = {},
  logger = null,
} = {}) {
  if (!page) {
    throw new Error("createDocmanUserGroup requires an existing Playwright page.");
  }

  const timeoutMs = Number(timeouts?.selectorMs) || 60000;
  const normalizedGroupName = String(groupName || "").trim();
  const requestedMembers = uniqueNames(usernames);

  if (!normalizedGroupName) {
    throw new Error("Group name is required to create a Docman user group.");
  }
  if (!requestedMembers.length) {
    throw new Error("At least one user is required to create a Docman user group.");
  }

  const stepToken = logger?.startStep?.("create_docman_user_group_internal", {
    groupName: normalizedGroupName,
    requestedMembers: requestedMembers.length,
  });

  try {
    await assertDocmanSessionReady(page, "creating user group");
    console.log("➡ Opening Docman User Groups…");

    await openUserGroupsList(page, timeoutMs);
    await startCreateGroup(page, timeoutMs);
    await fillGroupDescription(page, normalizedGroupName, timeoutMs);
    await openManageMembers(page, timeoutMs);

    const addedMembers = [];

    for (const memberName of requestedMembers) {
      const selectedName = await addMemberFromModal(page, memberName, timeoutMs);
      addedMembers.push(selectedName);
      logger?.event?.("user_group_member_added", {
        groupName: normalizedGroupName,
        requestedName: memberName,
        selectedName,
      });
    }

    await confirmMemberSelection(page, timeoutMs);
    await submitUserGroupCreate(page, normalizedGroupName, timeoutMs);

    const result = {
      groupName: normalizedGroupName,
      members: addedMembers,
      requestedMembers: requestedMembers.length,
      addedMembers: addedMembers.length,
    };

    logger?.endStep?.(stepToken, {
      status: "ok",
      groupName: normalizedGroupName,
      requestedMembers: requestedMembers.length,
      addedMembers: addedMembers.length,
    });

    return result;
  } catch (error) {
    logger?.endStep?.(stepToken, {
      status: "error",
      groupName: normalizedGroupName,
      requestedMembers: requestedMembers.length,
      errorMessage: error?.message || String(error),
    });

    const debug = await captureCreateGroupDebug(page, "failed");
    const message = error?.message || String(error);
    throw new Error(
      `${message}. Saved debug files: ${debug.screenshotPath}, ${debug.htmlPath}`
    );
  }
}

async function openUserGroupsList(page, timeoutMs) {
  await assertDocmanSessionReady(page, "opening User Groups");

  const directUrl = resolveDocmanUrl(page, USER_GROUPS_LIST_PATH);
  if (directUrl) {
    await page
      .goto(directUrl, {
        waitUntil: "domcontentloaded",
        timeout: Math.min(timeoutMs, 25000),
      })
      .catch(() => {});
  }

  await assertDocmanSessionReady(page, "opening User Groups");

  if (!(await isAnyVisible(page, USER_GROUP_LIST_SCREEN_SELECTORS, 2500))) {
    await clickAny(page, USER_GROUPS_MENU_SELECTORS, {
      timeoutMs: Math.min(timeoutMs, 15000),
      description: "User Groups menu",
      throwOnMissing: false,
    });
  }

  await waitAny(page, USER_GROUP_LIST_SCREEN_SELECTORS, {
    timeoutMs,
    description: "User Groups list screen",
  });
}

async function startCreateGroup(page, timeoutMs) {
  await clickAny(page, LIST_CREATE_BUTTON_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 15000),
    description: "Create user group button",
  });

  await waitAny(page, CREATE_GROUP_SCREEN_SELECTORS, {
    timeoutMs,
    description: "Create user group screen",
  });
}

async function fillGroupDescription(page, groupName, timeoutMs) {
  const input = await waitAny(page, GROUP_DESCRIPTION_INPUT_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 15000),
    description: "Group description input",
  });

  await input.locator.click({ clickCount: 3 }).catch(() => {});
  await input.locator.fill("").catch(() => {});
  await input.locator.type(groupName, { delay: 20 }).catch(async () => {
    await input.locator.fill(groupName).catch(() => {});
  });
  await input.locator.evaluate((element) => {
    element.dispatchEvent(new Event("input", { bubbles: true }));
    element.dispatchEvent(new Event("change", { bubbles: true }));
  }).catch(() => {});

  const typed = String((await input.locator.inputValue().catch(() => "")) || "").trim();
  if (typed !== groupName) {
    throw new Error(`Group description input did not keep value "${groupName}"`);
  }
}

async function openManageMembers(page, timeoutMs) {
  await clickAny(page, MANAGE_MEMBERS_BUTTON_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 15000),
    description: "Manage Members button",
  });

  console.log("➡ Opening Group Members selector…");
  await waitForDialog(page, /User Selection/i, timeoutMs);
}

async function addMemberFromModal(page, memberName, timeoutMs) {
  const modal = await waitForDialog(page, /User Selection/i, timeoutMs);
  console.log(`➡ Searching member: ${memberName}`);
  const searchInput = await ensureUserSearchTabActive(page, modal, timeoutMs);

  const searchTerm = resolveModalSearchTerm(memberName);
  await searchInput.click({ clickCount: 3 }).catch(() => {});
  await searchInput.fill("").catch(() => {});
  await searchInput.type(searchTerm, { delay: 30 }).catch(async () => {
    await searchInput.fill(searchTerm).catch(() => {});
  });
  await searchInput.evaluate((element) => {
    element.dispatchEvent(new Event("input", { bubbles: true }));
    element.dispatchEvent(new Event("change", { bubbles: true }));
  }).catch(() => {});

  const selectedAlready = await findBestSelectedName(modal, memberName);
  if (selectedAlready) {
    console.log(`↳ Already selected: ${selectedAlready}`);
    return selectedAlready;
  }

  const choice = await waitForBestModalCandidate(modal, memberName, timeoutMs);
  if (!choice) {
    const suggestions = await collectResultModalNames(modal);
    const suggestionMessage = suggestions.length
      ? ` Visible matches: ${suggestions.slice(0, 8).join(", ")}`
      : "";
    throw new Error(
      `Could not find a safe user match for "${memberName}" in User Search.${suggestionMessage}`
    );
  }

  const clicked = await clickExactTextInModalPane(modal, choice, "left");
  if (!clicked) {
    throw new Error(`Matched "${memberName}" to "${choice}" but could not click the result`);
  }

  const selected = await waitForModalSelection(modal, memberName, choice, timeoutMs);
  if (!selected) {
    throw new Error(`User "${choice}" was clicked but did not appear in Selection`);
  }

  console.log(`↳ Added member: ${memberName} -> ${selected}`);
  return selected;
}

async function confirmMemberSelection(page, timeoutMs) {
  const modal = await waitForDialog(page, /User Selection/i, timeoutMs);
  await clickWithin(modal, MODAL_CONFIRM_BUTTON_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 10000),
    description: "User Selection confirm button",
  });

  await waitForDialogClose(page, /User Selection/i, timeoutMs);
  await waitAny(page, CREATE_GROUP_SCREEN_SELECTORS, {
    timeoutMs,
    description: "Create user group screen after member selection",
  });
}

async function submitUserGroupCreate(page, groupName, timeoutMs) {
  await clickAny(page, PAGE_CREATE_BUTTON_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 15000),
    description: "Create group button",
  });

  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    await assertDocmanSessionReady(page, "finishing user group creation");

    const onList =
      /RecipientGroupList/i.test(String(page.url() || "")) ||
      (await isAnyVisible(page, USER_GROUP_LIST_SCREEN_SELECTORS, 600));
    if (onList) {
      return;
    }

    const bodyText = await readAllScopeText(page);
    const errorLine = extractRelevantErrorLine(bodyText);
    if (errorLine) {
      throw new Error(`Docman rejected user group creation: ${errorLine}`);
    }

    await sleep(200);
  }

  throw new Error(`Timed out waiting for Docman to finish creating group "${groupName}"`);
}

function resolveModalSearchTerm(memberName) {
  const value = String(memberName || "").trim();
  if (value.length >= 3) return value;
  throw new Error(
    `Cannot search for "${memberName}" because Docman User Search requires at least 3 characters`
  );
}

function extractRelevantErrorLine(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  return (
    lines.find((line) => /already exists/i.test(line)) ||
    lines.find((line) => /must select at least one user/i.test(line)) ||
    lines.find((line) => /required/i.test(line) && /description/i.test(line)) ||
    ""
  );
}

async function assertDocmanSessionReady(page, reason = "running Docman task") {
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

async function isLikelyDocmanLoginPage(page) {
  const url = String(page.url() || "");
  if (/\/account\/login/i.test(url) || /\/account\/prelogin/i.test(url)) {
    return true;
  }
  return isAnyVisible(page, DOCMAN_LOGIN_PAGE_SELECTORS, 500);
}

function resolveDocmanUrl(page, pathname) {
  const currentUrl = String(page.url() || "");
  try {
    const parsed = new URL(currentUrl);
    if (!String(parsed.host || "").toLowerCase().includes(DOCMAN_HOST_MATCH)) {
      return "";
    }
    return new URL(pathname, parsed.origin).toString();
  } catch (_) {
    return "";
  }
}

async function findVisibleDialog(page, labelPattern) {
  const explicitModal = page.locator("#usersearch_modal").first();
  const explicitVisible = await explicitModal.isVisible({ timeout: 150 }).catch(() => false);
  if (explicitVisible) {
    const text = await explicitModal.innerText().catch(() => "");
    if (!labelPattern || labelPattern.test(String(text || ""))) {
      return explicitModal;
    }
  }

  const rootSelectors = [
    "#usersearch_modal",
    "#usersearch",
    ".modal-dialog",
    ".modal-content",
    ".ui-dialog",
    "[role='dialog']",
    ".modal",
  ];

  for (const scope of getScopes(page)) {
    for (const rootSelector of rootSelectors) {
      const candidates = scope.locator(rootSelector).filter({ hasText: labelPattern });
      const count = Math.min(await candidates.count().catch(() => 0), 6);
      for (let i = 0; i < count; i += 1) {
        const locator = candidates.nth(i);
        const visible = await locator.isVisible({ timeout: 150 }).catch(() => false);
        if (visible) return locator;
      }
    }
  }

  return null;
}

async function waitForDialog(page, labelPattern, timeoutMs = 15000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const modal = await findVisibleDialog(page, labelPattern);
    if (modal) return modal;
    await sleep(200);
  }
  throw new Error(`Dialog not visible: ${labelPattern}`);
}

async function waitForDialogClose(page, labelPattern, timeoutMs = 15000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const modal = await findVisibleDialog(page, labelPattern);
    if (!modal) return;
    await sleep(200);
  }
  throw new Error(`Dialog did not close: ${labelPattern}`);
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
        const count = Math.min(await candidates.count().catch(() => 0), 8);
        for (let i = 0; i < count; i += 1) {
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

async function clickWithin(root, selectors, options = {}) {
  const { timeoutMs = 10000, description = "element" } = options;
  const locator = await waitWithin(root, selectors, { timeoutMs, description });
  await locator.scrollIntoViewIfNeeded().catch(() => {});
  await locator.click({ timeout: 5000 }).catch(async () => {
    await locator.click({ timeout: 5000, force: true }).catch(() => {});
  });
  await sleep(120);
  return locator;
}

async function waitWithin(root, selectors, options = {}) {
  const { timeoutMs = 10000, description = "element" } = options;
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    for (const selector of selectors) {
      const candidates = root.locator(selector);
      const count = Math.min(await candidates.count().catch(() => 0), 8);
      for (let i = 0; i < count; i += 1) {
        const locator = candidates.nth(i);
        const visible = await locator.isVisible({ timeout: 150 }).catch(() => false);
        if (visible) return locator;
      }
    }
    await sleep(150);
  }

  throw new Error(`${description} not visible inside dialog`);
}

async function ensureUserSearchTabActive(page, modal, timeoutMs) {
  const quickInput = await waitWithinOptional(modal, ["#SearchTerm"], 800);
  if (quickInput) {
    return quickInput;
  }

  const tab = await waitWithin(modal, USER_SEARCH_TAB_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 10000),
    description: "User Search tab",
  });

  await tab.scrollIntoViewIfNeeded().catch(() => {});
  await tab.click({ timeout: 5000 }).catch(async () => {
    await tab.click({ timeout: 5000, force: true }).catch(() => {});
  });

  await page.evaluate(() => {
    const anchor = document.querySelector('#usersearch_modal a[href="#tab2"]');
    if (anchor && typeof anchor.click === "function") {
      anchor.click();
    }
  }).catch(() => {});

  const input = await waitWithin(modal, MODAL_SEARCH_INPUT_SELECTORS, {
    timeoutMs: Math.min(timeoutMs, 10000),
    description: "User Search input",
  });

  await page.waitForFunction(
    () => {
      const tab = document.querySelector("#tab2");
      const input = document.querySelector("#SearchTerm");
      if (!tab || !input) return false;
      const tabVisible = tab.classList.contains("active") || tab.offsetParent !== null;
      const inputVisible = input.offsetParent !== null;
      return tabVisible && inputVisible;
    },
    { timeout: Math.min(timeoutMs, 10000) }
  ).catch(() => {});

  return input;
}

async function waitWithinOptional(root, selectors, timeoutMs = 1000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    for (const selector of selectors) {
      const locator = root.locator(selector).first();
      const visible = await locator.isVisible({ timeout: 150 }).catch(() => false);
      if (visible) return locator;
    }
    await sleep(100);
  }
  return null;
}

async function waitForExactResultClick(modal, memberName, timeoutMs) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const clicked = await clickExactTextInModalPane(modal, memberName, "left");
    if (clicked) return true;
    await sleep(200);
  }
  return false;
}

async function waitForBestModalCandidate(modal, memberName, timeoutMs) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const suggestions = await collectResultModalNames(modal);
    const resolved = resolveBestModalCandidateText(memberName, suggestions);
    if (resolved) return resolved;
    await sleep(200);
  }
  return "";
}

async function waitForModalSelection(modal, memberName, choice, timeoutMs) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const selected =
      (await findBestSelectedName(modal, choice)) ||
      (await findBestSelectedName(modal, memberName));
    if (selected) return selected;
    await sleep(200);
  }
  return "";
}

async function clickExactTextInModalPane(modal, text, side) {
  const exact = modal.locator(`xpath=.//*[normalize-space()=${toXPathLiteral(text)}]`);
  const modalBox = await modal.boundingBox().catch(() => null);
  const count = Math.min(await exact.count().catch(() => 0), 20);

  for (let i = 0; i < count; i += 1) {
    const locator = exact.nth(i);
    const visible = await locator.isVisible({ timeout: 150 }).catch(() => false);
    if (!visible) continue;

    const box = await locator.boundingBox().catch(() => null);
    if (!isBoxInModalPane(box, modalBox, side)) continue;

    await locator.scrollIntoViewIfNeeded().catch(() => {});
    const clicked =
      (await locator.click({ timeout: 1000 }).then(() => true).catch(() => false)) ||
      (await locator
        .locator("xpath=ancestor-or-self::*[self::a or self::button or self::tr or self::td or self::li or @role='button'][1]")
        .first()
        .click({ timeout: 1000 })
        .then(() => true)
        .catch(() => false));

    if (clicked) {
      await sleep(120);
      return true;
    }
  }

  return false;
}

async function isExactTextVisibleInModalPane(modal, text, side) {
  const exact = modal.locator(`xpath=.//*[normalize-space()=${toXPathLiteral(text)}]`);
  const modalBox = await modal.boundingBox().catch(() => null);
  const count = Math.min(await exact.count().catch(() => 0), 20);

  for (let i = 0; i < count; i += 1) {
    const locator = exact.nth(i);
    const visible = await locator.isVisible({ timeout: 150 }).catch(() => false);
    if (!visible) continue;

    const box = await locator.boundingBox().catch(() => null);
    if (isBoxInModalPane(box, modalBox, side)) {
      return true;
    }
  }

  return false;
}

async function findBestMatchingTextInModalPane(modal, text, side) {
  const visibleTexts = await collectModalPaneTexts(modal, side);
  return resolveBestModalCandidateText(text, visibleTexts);
}

async function findBestSelectedName(modal, text) {
  const selectedNames = await collectSelectedModalNames(modal);
  return resolveBestModalCandidateText(text, selectedNames);
}

async function collectResultModalNames(modal) {
  return collectModalListNames(modal, "results_list");
}

async function collectSelectedModalNames(modal) {
  return collectModalListNames(modal, "selected_list");
}

async function collectModalListNames(modal, listId) {
  const names = await modal
    .evaluate((element, targetListId) => {
      const normalize = (value) => String(value || "").replace(/\s+/g, " ").trim();
      const list = element.querySelector(`#${targetListId}`);
      if (!list) return [];

      const out = [];
      for (const item of Array.from(list.querySelectorAll("li"))) {
        const raw =
          item.getAttribute("data-name") ||
          item.querySelector(".name")?.getAttribute("title") ||
          item.querySelector(".name")?.textContent ||
          "";
        const text = normalize(raw);
        if (text && !out.includes(text)) {
          out.push(text);
        }
      }

      return out;
    }, listId)
    .catch(() => []);

  return Array.isArray(names) ? names : [];
}

async function collectModalPaneTexts(modal, side) {
  const modalBox = await modal.boundingBox().catch(() => null);
  if (!modalBox) return [];

  const texts = await modal
    .evaluate((element, desiredSide) => {
      const normalize = (value) => String(value || "").replace(/\s+/g, " ").trim();
      const isVisible = (node) => {
        if (!node) return false;
        const style = window.getComputedStyle(node);
        if (style.visibility === "hidden" || style.display === "none") return false;
        const rect = node.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      };

      const modalRect = element.getBoundingClientRect();
      const threshold = modalRect.left + modalRect.width * 0.52;
      const nodes = Array.from(
        element.querySelectorAll("li, td, a, button, div, span, label")
      );
      const out = [];

      for (const node of nodes) {
        if (!isVisible(node)) continue;
        const rect = node.getBoundingClientRect();
        const onLeft = rect.left < threshold;
        if (desiredSide === "left" && !onLeft) continue;
        if (desiredSide === "right" && onLeft) continue;

        const text = normalize(node.textContent);
        if (!text || text.length > 80) continue;
        if (/^confirm$|^cancel$|^user search$|^favourites$|^selection$/i.test(text)) continue;
        if (!out.includes(text)) out.push(text);
      }

      return out.slice(0, 20);
    }, side)
    .catch(() => []);

  return Array.isArray(texts) ? texts : [];
}

function resolveBestModalCandidateText(targetName, candidates = []) {
  const uniqueCandidates = uniqueNames(candidates);
  if (!uniqueCandidates.length) return "";

  const targetRaw = String(targetName || "").trim();
  if (!targetRaw) return "";

  const ranked = uniqueCandidates
    .map((candidate) => ({
      candidate,
      score: scoreCandidateMatch(targetRaw, candidate),
    }))
    .filter((row) => row.score > 0)
    .sort((a, b) => b.score - a.score || a.candidate.localeCompare(b.candidate));

  if (!ranked.length) return "";

  const best = ranked[0];
  const ties = ranked.filter((row) => row.score === best.score);
  if (ties.length > 1) {
    return "";
  }

  return best.candidate;
}

function scoreCandidateMatch(targetName, candidateName) {
  const rawTarget = normalizeWhitespace(targetName).toLowerCase();
  const rawCandidate = normalizeWhitespace(candidateName).toLowerCase();
  const target = normalizeUserNameForMatching(targetName);
  const candidate = normalizeUserNameForMatching(candidateName);

  if (!target || !candidate) return 0;
  if (rawTarget === rawCandidate) return 100;
  if (target === candidate) return 95;

  const targetTokens = target.split(" ").filter(Boolean);
  const candidateTokens = candidate.split(" ").filter(Boolean);
  if (!targetTokens.length || !candidateTokens.length) return 0;

  const allTokensMatch = targetTokens.every((token) => candidateTokens.includes(token));
  if (!allTokensMatch) return 0;

  if (candidate.endsWith(target) || candidate.startsWith(target)) {
    return 90;
  }

  if (candidate.includes(target)) {
    return 85;
  }

  return 70 - Math.max(0, candidateTokens.length - targetTokens.length);
}

function normalizeUserNameForMatching(value) {
  return normalizeWhitespace(
    String(value || "")
      .toLowerCase()
      .replace(/\b(mr|mrs|miss|ms|dr|prof|professor|sir|lady)\b/g, " ")
      .replace(/[^a-z0-9\s]/g, " ")
  );
}

function normalizeWhitespace(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function isBoxInModalPane(box, modalBox, side) {
  if (!box || !modalBox) return true;
  const threshold = modalBox.x + modalBox.width * 0.52;
  const isLeft = box.x < threshold;
  return side === "left" ? isLeft : !isLeft;
}

function getScopes(page) {
  const scopes = [page];
  for (const frame of page.frames()) {
    if (frame === page.mainFrame()) continue;
    scopes.push(frame);
  }
  return scopes;
}

async function readAllScopeText(page) {
  const chunks = [];
  for (const scope of getScopes(page)) {
    const text = await scope.locator("body").innerText().catch(() => "");
    if (text) chunks.push(text);
  }
  return chunks.join("\n");
}

function uniqueNames(values = []) {
  const seen = new Set();
  const out = [];

  for (const value of values) {
    const text = String(value || "").trim();
    if (!text) continue;
    const key = text.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(text);
  }

  return out;
}

function toXPathLiteral(value) {
  const raw = String(value || "");
  if (!raw.includes("'")) return `'${raw}'`;
  if (!raw.includes('"')) return `"${raw}"`;
  return `concat(${raw.split("'").map((part) => `'${part}'`).join(`, "'", `)})`;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function captureCreateGroupDebug(page, suffix = "debug") {
  const safe = String(suffix || "debug").replace(/[^a-z0-9_-]+/gi, "-");
  const screenshotPath = `create-user-group-${safe}.png`;
  const htmlPath = `create-user-group-${safe}.html`;

  await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
  const html = await page.content().catch(() => "");
  if (html) {
    fs.writeFileSync(htmlPath, html, "utf8");
  }

  return { screenshotPath, htmlPath };
}

module.exports = createDocmanUserGroup;
