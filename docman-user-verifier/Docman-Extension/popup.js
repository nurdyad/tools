const practiceInput = document.getElementById("practice");
const runBtn = document.getElementById("run");
const statusEl = document.getElementById("status");

async function refreshStatus() {
  try {
    const response = await chrome.runtime.sendMessage({ type: "status" });
    if (response?.status) statusEl.textContent = response.status;
  } catch {
    statusEl.textContent = "Ready";
  }
}

runBtn.addEventListener("click", async () => {
  const query = practiceInput.value.trim();
  if (!query) {
    statusEl.textContent = "Enter a practice name.";
    return;
  }

  statusEl.textContent = "Working…";
  await chrome.runtime.sendMessage({ type: "startLookup", query });
  await refreshStatus();
});

refreshStatus();
