// run.js
const inquirer = require("inquirer").default;
const clipboardy = require("clipboardy").default;

const bootstrapDocmanSession = require("./automation/bootstrapDocmanSession");
const verifyDocmanUsers = require("./verifyDocmanUsers");
const cleanBetterLetterProcessing = require("./cleanBetterLetterProcessing");

(async () => {
  let session = null;

  try {
    const { practiceName } = await inquirer.prompt([
      {
        type: "input",
        name: "practiceName",
        message: "Enter Practice Name (as shown in BetterLetter):",
      },
    ]);

    if (!practiceName?.trim()) {
      console.log("Cancelled.");
      return;
    }

    console.log("🔗 Bootstrapping BetterLetter → Docman session…");

    // ✅ Option B: hardcode HTTP Basic Auth here (browser-level popup)
    session = await bootstrapDocmanSession(practiceName.trim(), {
      httpCredentials: {
        username: "mailroom_admin",
        password: "Yxbq95wbYhp0sAbR8xmV",
      },
    });

    const { page } = session;

    const { mode } = await inquirer.prompt([
      {
        type: "list",
        name: "mode",
        message: "What would you like to do?",
        choices: [
          { name: "Verify Docman users (copy valid names)", value: "verify" },
          { name: "Clean folder (move NON-UUID documents)", value: "clean" },
        ],
      },
    ]);

    if (mode === "verify") {
      const { usernamesRaw } = await inquirer.prompt([
        {
          type: "editor",
          name: "usernamesRaw",
          message: "Paste Docman usernames to verify (one per line):",
        },
      ]);

      const usernames = (usernamesRaw || "")
        .split("\n")
        .map((u) => u.trim())
        .filter(Boolean);

      if (!usernames.length) {
        console.log("No usernames provided.");
        return;
      }

      const results = await verifyDocmanUsers({ page, usernames });

      console.log("\nVerification results:");
      console.table(results);

      const valid = results
        .filter((r) => r.exists && r.docmanUsername)
        .map((r) => r.docmanUsername);

      if (valid.length) {
        const output = valid.join("\n");
        await clipboardy.write(output);
        console.log("\n========================================");
        console.log("READY FOR BETTERLETTER EXTENSION (exact matches only)");
        console.log("========================================\n");
        console.log(output);
        console.log("\n========================================");
        console.log("Copied to clipboard ✔");
      } else {
        console.log("\nNo valid Docman users found.");
      }

      return;
    }

    if (mode === "clean") {
      await cleanBetterLetterProcessing({
        page,
        batchSize: 50,
        dryRun: false,
      });
      console.log("\n✔ Done.");
      return;
    }
  } catch (err) {
    console.error("\n❌ FAILED:", err?.message || err);
  } finally {
    try {
      if (session?.context) {
        await session.context.close();
      }
    } catch (_) {}

    process.stdin.pause();
  }
})();
