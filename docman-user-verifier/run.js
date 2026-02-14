// run.js
const inquirer = require("inquirer").default;
const clipboardy = require("clipboardy").default;

const bootstrapDocmanSession = require("./automation/bootstrapDocmanSession");
const verifyDocmanUsers = require("./verifyDocmanUsers");
const cleanBetterLetterProcessing = require("./cleanBetterLetterProcessing");

(async () => {
  let session = null;
  let sessionCloseMessage = "Session finished. Browser will stay open.";

  try {
    printPhaseBanner("Step 0", ["Collect target practice"]);
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

    const { mode: modeInput } = await inquirer.prompt([
      {
        type: "list",
        name: "mode",
        message: "What would you like to do?",
        default: "login",
        choices: [
          { name: "Just login to Docman (no task)", value: "login" },
          { name: "Verify Docman users (copy valid names)", value: "verify" },
          { name: "Clean folder (move NON-UUID documents)", value: "clean" },
        ],
      },
    ]);

    const mode = normalizeMode(modeInput);
    if (!mode) {
      console.log("Unknown mode selected:", modeInput);
      return;
    }

    console.log(`✅ Mode selected: ${mode}`);
    console.log("Using basic auth user:", "mailroom_admin");

    console.log("🔗 Bootstrapping BetterLetter → Docman session…");

    // ✅ HTTP Basic Auth hardcoded here (browser popup layer)
    session = await bootstrapDocmanSession(practiceName.trim(), {
      httpCredentials: {
        username: "mailroom_admin",
        password: "Yxbq95wbYhp0sAbR8xmV",
      },
      forceFreshDocmanLogin: false,
      resetDocmanAuthAtStart: true,
      includeDocmanInHealthCheck: false,
      skipPostLoginDialogWatch: mode === "login",
    });

    const { page } = session;

    if (mode === "login") {
      sessionCloseMessage = "Login Successful. Browser will stay open.";
      return;
    }

    if (mode === "clean") {
      const { confirmClean } = await inquirer.prompt([
        {
          type: "confirm",
          name: "confirmClean",
          default: false,
          message:
            "CLEAN mode moves documents between folders. Continue with CLEAN?",
        },
      ]);

      if (!confirmClean) {
        console.log("Cancelled CLEAN workflow.");
        return;
      }

      if (typeof bootstrapDocmanSession.gotoDocmanFilingAndActivate === "function") {
        console.log("➡ Preparing Docman Filing for CLEAN workflow…");
        await bootstrapDocmanSession.gotoDocmanFilingAndActivate(page, {
          skipDialogCheck: true,
        });
      }

      console.log("🧹 Starting CLEAN workflow…");
      await cleanBetterLetterProcessing({
        page,
        batchSize: 50,
        dryRun: false,
      });
      console.log("✅ CLEAN workflow finished.");
      return;
    }

    if (mode === "verify") {
      console.log("🔍 Starting VERIFY workflow…");

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

      console.log("✅ VERIFY workflow finished.");
      return;
    }
  } catch (err) {
    console.error("\n❌ FAILED:", err?.message || err);
  } finally {
    if (session?.context) {
      await waitForManualSessionEnd(session.context, sessionCloseMessage);
    } else {
      process.stdin.pause();
    }
  }
})();

async function waitForManualSessionEnd(
  context,
  sessionMessage = "Session finished. Browser will stay open."
) {
  console.log(`\n${sessionMessage}`);
  console.log("Close the browser window, or press ENTER here to close it and exit.");

  await new Promise((resolve) => {
    process.stdin.resume();
    process.stdin.once("data", () => resolve());
  });

  try {
    await context.close();
  } catch (_) {}

  process.stdin.pause();
}

function printPhaseBanner(title, lines = []) {
  const width = 62;
  const bar = "=".repeat(width);
  console.log(`\n${bar}`);
  console.log(` ${title}`);
  for (const line of lines) {
    console.log(` - ${line}`);
  }
  console.log(`${bar}\n`);
}

function normalizeMode(modeInput) {
  const mode = String(modeInput || "").trim().toLowerCase();
  if (mode === "login" || mode === "login_only") return "login";
  if (mode === "clean") return "clean";
  if (mode === "verify") return "verify";
  return null;
}
