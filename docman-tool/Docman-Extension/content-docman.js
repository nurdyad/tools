chrome.runtime.onMessage.addListener((message) => {
  if (message?.type !== "fillDocman") return;

  const { odsCode, username, password } = message.payload || {};
  if (!odsCode || !username || !password) return;

  const odsInput = document.querySelector("#OdsCode");
  const userInput = document.querySelector("#Username");
  const passInput = document.querySelector("#Password");
  const submitBtn = document.querySelector("button[type='submit']");

  if (!odsInput || !userInput || !passInput || !submitBtn) return;

  odsInput.focus();
  odsInput.value = odsCode;
  odsInput.dispatchEvent(new Event("input", { bubbles: true }));

  userInput.focus();
  userInput.value = username;
  userInput.dispatchEvent(new Event("input", { bubbles: true }));

  passInput.focus();
  passInput.value = password;
  passInput.dispatchEvent(new Event("input", { bubbles: true }));

  submitBtn.click();
});
