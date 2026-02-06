const inquirer = require("inquirer").default;
const clipboardy = require("clipboardy").default;
const verifyDocmanUsers = require("./verifyDocmanUsers");

(async () => {
  const answers = await inquirer.prompt([
    {
      type: "input",
      name: "odsCode",
      message: "Enter ODS code:"
    },
    {
      type: "input",
      name: "adminUsername",
      message: "Enter Docman admin username:"
    },
    {
      type: "password",
      name: "adminPassword",
      message: "Enter Docman admin password:",
      mask: "*"
    },
    {
      type: "editor",
      name: "usernames",
      message: "Paste Docman usernames to verify (one per line):"
    }
  ]);

  const usernames = answers.usernames
    .split("\n")
    .map(u => u.trim())
    .filter(Boolean);

  const results = await verifyDocmanUsers({
    odsCode: answers.odsCode,
    adminUsername: answers.adminUsername,
    adminPassword: answers.adminPassword,
    usernames
  });

  console.log("\nVerification results:");
  console.table(results);

  // ✅ Collect ONLY valid Docman usernames (exact, unchanged)
  const validDocmanUsers = results
    .filter(r => r.exists && r.docmanUsername)
    .map(r => r.docmanUsername);

  if (!validDocmanUsers.length) {
    console.log("\nNo valid Docman users found.");
    return;
  }

  const outputBlock = validDocmanUsers.join("\n");

  console.log("\n========================================");
  console.log("READY FOR BETTERLETTER EXTENSION");
  console.log("(copy below)");
  console.log("========================================\n");
  console.log(outputBlock);
  console.log("\n========================================");

  await clipboardy.write(outputBlock);
  console.log("Copied to clipboard ✔");
  console.log("========================================");
})();
