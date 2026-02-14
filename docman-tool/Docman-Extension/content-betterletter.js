const PENDING_QUERY_KEY = "pendingQuery";

// Marker for quick debugging from page console
document.documentElement.setAttribute("data-docman-extension", "loaded");

chrome.runtime.onMessage.addListener((message) => {
  if (message?.type === "runLookup") {
    const query = (message.query || "").trim();
    if (query) {
      runWithQuery(query);
      return;
    }
    runIfPending();
  }
});

setStatus("Extension loaded on BetterLetter.");
runIfPending();

async function runIfPending() {
  const data = await chrome.storage.local.get(PENDING_QUERY_KEY);
  const query = (data[PENDING_QUERY_KEY] || "").trim();
  if (!query) return;

  await runWithQuery(query);
}

async function runWithQuery(query) {
  const path = window.location.pathname;
  if (path === "/admin_panel/practices") {
    await findPracticeAndNavigate(query);
  } else if (path.startsWith("/admin_panel/practices/")) {
    await extractCredsFromPracticePage(query);
  }
}

function setStatus(message) {
  chrome.storage.local.set({ status: message });
}

async function findPracticeAndNavigate(query) {
  setStatus(`Looking for practice matching: ${query}`);

  // If not logged in, BetterLetter redirects to login
  if (document.querySelector("form[action*='login'], form#new_user")) {
    setStatus("Please log into BetterLetter, then try again.");
    return;
  }

  try {
    await waitFor(() => document.querySelector("table tbody tr"), 20000);
  } catch {
    setStatus("No practices table found. Waiting for table…");
  }

  let rows = Array.from(document.querySelectorAll("table tbody tr"));
  if (!rows.length) {
    await waitForMutation("table tbody", 15000);
    rows = Array.from(document.querySelectorAll("table tbody tr"));
  }

  if (!rows.length) {
    setStatus("No practices rows found.");
    return;
  }

  setStatus(`Found ${rows.length} rows. Searching…`);

  const normalizedQuery = query.toLowerCase();

  let matchedRow = null;
  for (const row of rows) {
    const rowText = row.innerText?.toLowerCase() || "";
    if (rowText.includes(normalizedQuery)) {
      matchedRow = row;
      break;
    }
  }

  if (!matchedRow) {
    setStatus(`Practice not found: ${query}`);
    return;
  }

  const link = matchedRow.querySelector("a[href*='/admin_panel/practices/']");
  if (link) {
    setStatus("Opening practice details…");
    link.click();
    return;
  }

  const odsCode = extractOdsFromRow(matchedRow);
  if (!odsCode) {
    setStatus("Could not read ODS code from the list.");
    return;
  }

  setStatus(`Found ODS ${odsCode}. Opening practice details…`);
  window.location.href = `/admin_panel/practices/${odsCode}`;
}

async function extractCredsFromPracticePage(query) {
  setStatus("Opening EHR Settings…");

  const ehrTabSelector =
    "[data-test-id='tab-ehr_settings'], [phx-value-tab='ehr_settings']";
  try {
    await waitFor(() => document.querySelector(ehrTabSelector), 20000);
  } catch {
    setStatus("EHR Settings tab not found.");
    return;
  }

  const ehrTab = document.querySelector(ehrTabSelector);
  if (ehrTab) {
    ehrTab.scrollIntoView({ behavior: "smooth", block: "center" });
    ehrTab.click();
  }

  const usernameInput = "#ehr_settings\\[docman\\]\\[username\\]";
  const passwordInput = "#ehr_settings\\[docman\\]\\[password\\]";

  await waitFor(() => document.querySelector(usernameInput), 20000);
  await waitFor(() => document.querySelector(passwordInput), 20000);

  const username = document.querySelector(usernameInput)?.value?.trim();
  const password = document.querySelector(passwordInput)?.value?.trim();

  let odsCode =
    document.querySelector("span.text-white.bg-subtle")?.innerText?.trim() ||
    window.location.pathname.split("/").pop();

  if (!username || !password || !odsCode) {
    setStatus("Missing Docman credentials in EHR Settings.");
    return;
  }

  setStatus(`Docman credentials found for ${odsCode}.`);
  await chrome.storage.local.remove(PENDING_QUERY_KEY);

  chrome.runtime.sendMessage({
    type: "docmanCreds",
    payload: { odsCode, username, password },
  });
}

function extractOdsFromRow(row) {
  const firstCell = row.querySelector("td");
  const text = (firstCell?.innerText || "").trim();
  const match = text.match(/[A-Z][0-9]{5}/);
  return match ? match[0] : null;
}

function waitFor(predicate, timeoutMs = 10000, intervalMs = 200) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const timer = setInterval(() => {
      if (predicate()) {
        clearInterval(timer);
        resolve(true);
      } else if (Date.now() - start > timeoutMs) {
        clearInterval(timer);
        reject(new Error("Timeout"));
      }
    }, intervalMs);
  });
}

function waitForMutation(selector, timeoutMs = 10000) {
  return new Promise((resolve, reject) => {
    const target = document.querySelector(selector);
    if (!target) {
      reject(new Error("Target not found"));
      return;
    }

    const start = Date.now();
    const observer = new MutationObserver(() => {
      if (Date.now() - start > timeoutMs) {
        observer.disconnect();
        reject(new Error("Timeout"));
      } else {
        observer.disconnect();
        resolve(true);
      }
    });

    observer.observe(target, { childList: true, subtree: true });

    setTimeout(() => {
      observer.disconnect();
      reject(new Error("Timeout"));
    }, timeoutMs);
  });
}
