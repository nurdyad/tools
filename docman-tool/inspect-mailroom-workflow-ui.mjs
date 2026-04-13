import { chromium } from "playwright";
import { writeFile } from "node:fs/promises";

const ods = process.argv[2] || "G82083";
const storageStatePath =
  "/Users/nursiddique/chrome-extensions/MailroomNavigator/.automation-state/storageState.mailroomnavigator.json";
const targetUrl = `https://app.betterletter.ai/admin_panel/practices/${ods}`;

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

const browser = await chromium.launch({ headless: true });
const context = await browser.newContext({ storageState: storageStatePath });
const page = await context.newPage();

const collectState = async () => page.evaluate(() => {
  const visible = (element) =>
    Boolean(element && (element.offsetParent !== null || element.getClientRects?.().length));

  const docmanInputs = Array.from(document.querySelectorAll("input[name='form-[docman_group]']")).filter(visible);
  const labelInputs = Array.from(document.querySelectorAll("input[name='form-[label_for_ui]']")).filter(visible);
  const tabsAndButtons = Array.from(document.querySelectorAll("button, a, [role='tab'], [role='button']"))
    .filter(visible)
    .map((element) => ({
      text: String(element.textContent || "").replace(/\s+/g, " ").trim(),
      disabled: Boolean(element.disabled),
      phxClick: String(element.getAttribute("phx-click") || ""),
      testId: String(element.getAttribute("data-test-id") || ""),
      role: String(element.getAttribute("role") || ""),
      href: String(element.getAttribute("href") || ""),
      className: String(element.className || ""),
    }));
  const buttons = Array.from(document.querySelectorAll("button, [role='button']")).filter(visible).map((button) => ({
    text: String(button.textContent || "").replace(/\s+/g, " ").trim(),
    disabled: Boolean(button.disabled),
    phxClick: String(button.getAttribute("phx-click") || ""),
    testId: String(button.getAttribute("data-test-id") || ""),
    className: String(button.className || ""),
  }));

  const lastDocman = docmanInputs.at(-1) || null;
  const lastLabel = labelInputs.at(-1) || null;
  const commonAncestor = (() => {
    if (!lastDocman || !lastLabel) return null;
    const visited = new Set();
    let current = lastDocman;
    while (current) {
      visited.add(current);
      current = current.parentElement;
    }
    current = lastLabel;
    while (current) {
      if (visited.has(current)) return current;
      current = current.parentElement;
    }
    return null;
  })();

  return {
    url: location.href,
    heading: document.querySelector("h1,h2,h3")?.textContent || "",
    tabsAndButtons,
    docmanInputs: docmanInputs.map((input) => ({
      name: input.name,
      value: input.value,
      placeholder: input.getAttribute("placeholder") || "",
      id: input.id || "",
      outerHTML: input.outerHTML,
    })),
    labelInputs: labelInputs.map((input) => ({
      name: input.name,
      value: input.value,
      placeholder: input.getAttribute("placeholder") || "",
      id: input.id || "",
      outerHTML: input.outerHTML,
    })),
    buttons,
    draftHtml: commonAncestor?.outerHTML || "",
    bodyText: document.body.innerText,
  };
});

let failureMessage = "";
try {
  await page.goto(targetUrl, { waitUntil: "domcontentloaded", timeout: 120000 });
  await page.waitForTimeout(2000);

  const taskRecipientsTab = page
    .locator("[data-test-id='tab-task_recipients']")
    .or(page.getByRole("tab", { name: /task recipients/i }))
    .or(page.getByRole("button", { name: /task recipients/i }))
    .or(page.getByRole("link", { name: /task recipients/i }))
    .first();

  await taskRecipientsTab.click({ timeout: 30000 });
  await page.waitForTimeout(1500);

  const addButton = page.locator("[phx-click='add_workflow_group'], [phx-click*='add_workflow']").first();
  await addButton.click({ timeout: 30000 });
  await page.waitForTimeout(1200);
} catch (error) {
  failureMessage = String(error?.message || error);
} finally {
  const state = await collectState();
  await page.screenshot({
    path: `/Users/nursiddique/tools/docman-tool/inspect-mailroom-workflow-ui-${ods}.png`,
    fullPage: true,
  });
  await writeFile(
    `/Users/nursiddique/tools/docman-tool/inspect-mailroom-workflow-ui-${ods}.json`,
    JSON.stringify(
      {
        ...state,
        heading: normalizeText(state.heading),
        bodyPreview: normalizeText(state.bodyText).slice(0, 4000),
        failureMessage,
      },
      null,
      2,
    ),
    "utf8",
  );
  await browser.close();
}

if (failureMessage) {
  throw new Error(failureMessage);
}
