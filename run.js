const verifyDocmanUsers = require("./verifyDocmanUsers");

(async () => {
  const inquirer = (await import("inquirer")).default;
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
      message:
        "Paste Docman usernames to verify (one per line):"
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
})();
