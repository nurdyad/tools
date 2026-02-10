const PRACTICES_URL = "https://app.betterletter.ai/admin_panel/practices";
const STATUS_KEY = "status";
const PENDING_QUERY_KEY = "pendingQuery";

async function setStatus(message) {
  await chrome.storage.local.set({ [STATUS_KEY]: message });
}

chrome.runtime.onInstalled.addListener(() => {
  setStatus("Ready");
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type === "startLookup") {
    const query = (message.query || "").trim();
    if (!query) {
      setStatus("Enter a practice name first.");
      sendResponse({ ok: false, error: "Missing practice name" });
      return true;
    }

    (async () => {
      await setStatus(`Searching for practice: ${query}`);
      await chrome.storage.local.set({ [PENDING_QUERY_KEY]: query });

      const tab = await ensurePracticesTab();
      if (tab?.id) {
        await chrome.tabs.update(tab.id, { active: true });
        // Ensure content script is injected, then run lookup
        try {
          await chrome.scripting.executeScript({
            target: { tabId: tab.id },
            files: ["content-betterletter.js"],
          });
        } catch {
          // ignore injection errors; content script may already be present
        }

        try {
          await chrome.tabs.sendMessage(tab.id, {
            type: "runLookup",
            query,
          });
        } catch {
          // content script may not be ready yet; it will run on load
        }
      }

      sendResponse({ ok: true });
    })();

    return true;
  }

  if (message?.type === "docmanCreds") {
    (async () => {
      const { odsCode, username, password } = message.payload || {};
      if (!odsCode || !username || !password) {
        await setStatus("Missing Docman creds from BetterLetter.");
        return;
      }

      await setStatus(`Opening Docman for ODS ${odsCode}…`);

      const tab = await chrome.tabs.create({
        url: "https://production.docman.thirdparty.nhs.uk/Account/Login",
        active: true,
      });

      const tabId = tab.id;
      if (!tabId) return;

      await waitForTabComplete(tabId);

      await chrome.tabs.sendMessage(tabId, {
        type: "fillDocman",
        payload: { odsCode, username, password },
      });

      await setStatus("Docman login submitted.");
    })();

    sendResponse({ ok: true });
    return true;
  }

  if (message?.type === "status") {
    (async () => {
      const data = await chrome.storage.local.get(STATUS_KEY);
      sendResponse({ status: data[STATUS_KEY] || "Ready" });
    })();
    return true;
  }
});

async function ensurePracticesTab() {
  const tabs = await chrome.tabs.query({ url: "https://app.betterletter.ai/*" });

  for (const tab of tabs) {
    if (tab.url && tab.url.startsWith(PRACTICES_URL)) {
      return tab;
    }
  }

  // No existing tab; create a new one
  return chrome.tabs.create({ url: PRACTICES_URL, active: true });
}

function waitForTabComplete(tabId, timeoutMs = 30000) {
  return new Promise((resolve) => {
    const start = Date.now();

    const listener = (updatedTabId, info) => {
      if (updatedTabId === tabId && info.status === "complete") {
        chrome.tabs.onUpdated.removeListener(listener);
        resolve(true);
      } else if (Date.now() - start > timeoutMs) {
        chrome.tabs.onUpdated.removeListener(listener);
        resolve(false);
      }
    };

    chrome.tabs.onUpdated.addListener(listener);
  });
}
