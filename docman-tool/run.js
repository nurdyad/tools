// run.js
const fs = require("fs");
const path = require("path");
const inquirer = require("inquirer").default;
const clipboardy = require("clipboardy").default;

const bootstrapDocmanSession = require("./automation/bootstrapDocmanSession");
const verifyDocmanUsers = require("./verifyDocmanUsers");
const cleanBetterLetterProcessing = require("./cleanBetterLetterProcessing");
const onboardingDocmanFolders = require("./automation/onboardingDocmanFolders");
const createDocmanUserGroup = require("./automation/createDocmanUserGroup");
const {
  loadBetterLetterBasicAuth,
  saveBetterLetterBasicAuth,
  AUTH_FILE: BASIC_AUTH_FILE,
} = require("./automation/betterLetterBasicAuth");
const { loadEnvFile } = require("./automation/env");
const { loadRuntimeConfig } = require("./automation/config");
const { withRetry } = require("./automation/retry");
const { classifyError, createRunLogger } = require("./automation/runLogger");
const { parseCliArgs, printCliHelp } = require("./automation/cliArgs");

(async () => {
  let session = null;
  let runOutcome = "success";
  let cliOptions = null;
  const noopLogger = createRunLogger({ enabled: false });
  let runLogger = noopLogger;

  try {
    const envInfo = loadEnvFile();
    const { config, configPath, hasConfigFile, projectRoot } = loadRuntimeConfig();

    try {
      cliOptions = parseCliArgs(process.argv.slice(2));
    } catch (error) {
      throw new Error(
        `${error?.message || "Invalid CLI options"}. Run \"node run.js --help\" for usage.`
      );
    }

    if (cliOptions.help) {
      printCliHelp();
      return;
    }

    runLogger = createRunLogger({
      enabled: config.logging.enabled,
      projectRoot,
      logDirectory: config.logging.directory,
      app: "docman-tool",
    });

    runLogger.event("config_loaded", {
      configPath,
      hasConfigFile,
      envFile: envInfo.filePath,
      envFileExists: envInfo.exists,
      envLoadedCount: envInfo.loadedCount,
    });

    runLogger.event("cli_options_resolved", {
      nonInteractive: cliOptions.nonInteractive,
      hasPracticeArg: Boolean(cliOptions.practiceName),
      hasModeArg: Boolean(cliOptions.mode),
      hasCleanTypeArg: Boolean(cliOptions.clean.type),
      hasCleanSourceArg: Boolean(cliOptions.clean.sourceFolder),
      hasCleanDestinationArg: Boolean(cliOptions.clean.destinationFolder),
      cleanAutoConfirm: Boolean(cliOptions.clean.autoConfirm),
      verifyCliUsernames: cliOptions.verify.usernames.length,
      hasVerifyFileArg: Boolean(cliOptions.verify.usernamesFile),
      hasGroupNameArg: Boolean(cliOptions.createGroup.groupName),
    });

    console.log(`ℹ Runtime config: ${path.basename(configPath)} (${hasConfigFile ? "loaded" : "defaults"})`);
    if (envInfo.exists) {
      console.log(`ℹ Environment: ${path.basename(envInfo.filePath)} loaded (${envInfo.loadedCount} variable(s))`);
    }
    if (runLogger.enabled && runLogger.filePath) {
      console.log(`ℹ Run log: ${runLogger.filePath}`);
    }

    printPhaseBanner("Step 0", ["Collect target practice"]);

    let practiceName = String(cliOptions.practiceName || "").trim();
    if (practiceName) {
      console.log(`✔ Enter Practice Name (as shown in BetterLetter): ${practiceName}`);
    } else {
      const answer = await inquirer.prompt([
        {
          type: "input",
          name: "practiceName",
          message: "Enter Practice Name (as shown in BetterLetter):",
        },
      ]);
      practiceName = String(answer.practiceName || "").trim();
    }

    if (!practiceName) {
      console.log("Cancelled.");
      runOutcome = "cancelled";
      return;
    }

    runLogger.event("practice_selected", {
      practiceName,
      source: cliOptions.practiceName ? "cli" : "prompt",
    });

    let modeInput = "";

    if (cliOptions.mode) {
      modeInput = String(cliOptions.mode || "").trim();
      console.log(`✔ What would you like to do? ${modeInput}`);
    } else if (cliOptions.nonInteractive) {
      modeInput = String(config.modeDefaults.defaultMode || "").trim();
      console.log(`✔ What would you like to do? ${modeInput} (from config default)`);
    } else {
      const answer = await inquirer.prompt([
        {
          type: "list",
          name: "mode",
          message: "What would you like to do?",
          default: config.modeDefaults.defaultMode,
          choices: getModeChoices(),
        },
      ]);
      modeInput = answer.mode;
    }

    const mode = normalizeMode(modeInput);
    if (!mode) {
      throw new Error(`Unknown mode selected: ${modeInput}`);
    }

    validateNonInteractiveMode(cliOptions, mode);

    runLogger.event("mode_selected", {
      mode,
      source: cliOptions.mode ? "cli" : cliOptions.nonInteractive ? "config_default" : "prompt",
    });

    let basicAuth = resolveBasicAuth(config);
    if (!basicAuth) {
      basicAuth = await promptForBasicAuth({
        logger: runLogger,
      });
    }

    console.log(`✅ Mode selected: ${mode}`);
    console.log("Using basic auth user:", basicAuth.username);

    runLogger.event("basic_auth_resolved", {
      source: basicAuth.source,
      username: basicAuth.username,
    });

    console.log("🔗 Bootstrapping BetterLetter → Docman session…");

    session = await runStepWithRetry({
      label: "bootstrap_session",
      logger: runLogger,
      retryPolicy: { attempts: 1, baseDelayMs: 0, maxDelayMs: 0 },
      task: async () =>
        bootstrapDocmanSession(practiceName, {
          httpCredentials: {
            username: basicAuth.username,
            password: basicAuth.password,
          },
          window: config.browser.window,
          preDocmanHeadless: config.browser.preDocmanHeadless,
          allowVisibleFallback: config.browser.allowVisibleFallback,
          step4Visible: config.browser.step4Visible,
          step4UseCurrentChrome: config.browser.step4UseCurrentChrome,
          step4BrowserEngine: config.browser.step4BrowserEngine,
          chromeCdpUrl: config.browser.chromeCdpUrl,
          forceFreshDocmanLogin: false,
          resetDocmanAuthAtStart: true,
          includeDocmanInHealthCheck: false,
          skipPostLoginDialogWatch: mode === "login",
          retryPolicy: config.retries.step,
          timeouts: config.timeouts,
          logger: runLogger,
        }),
    });

    const { page } = session;

    let activeMode = mode;
    let activeCliOptions = cliOptions;

    while (true) {
      const taskResult = await executeMode({
        mode: activeMode,
        page,
        session,
        cliOptions: activeCliOptions,
        config,
        runLogger,
      });

      runOutcome = taskResult.outcome;

      if (cliOptions.nonInteractive) {
        break;
      }

      if (isSessionClosed(session, page)) {
        console.log("\nSession ended because the browser window was closed.");
        break;
      }

      const runAnotherTask = await promptForAnotherTask(practiceName);
      if (!runAnotherTask || isSessionClosed(session, page)) {
        break;
      }

      activeCliOptions = buildFollowUpCliOptions(cliOptions);
      activeMode = await promptForModeSelection(config);
      if (!activeMode || isSessionClosed(session, page)) {
        break;
      }
    }
  } catch (err) {
    runOutcome = "failed";
    if (runLogger?.enabled) {
      runLogger.event("run_error", {
        errorType: classifyError(err),
        errorMessage: err?.message || String(err),
      });
    }
    console.error("\n❌ FAILED:", err?.message || err);
  } finally {
    if (runLogger?.enabled) {
      runLogger.close({ outcome: runOutcome });
    }

    if (session?.context) {
      const shouldKeepHeadedSessionOpen =
        runOutcome === "failed" &&
        !cliOptions?.nonInteractive &&
        !session?.headless;

      if (shouldKeepHeadedSessionOpen) {
        await waitForManualSessionEnd(
          session,
          "Run failed. Browser will stay open for inspection."
        );
      } else {
        await closeSession(session);
      }
    } else {
      process.stdin.pause();
    }
  }
})();

async function runStepWithRetry({ label, task, logger, retryPolicy }) {
  const stepToken = logger?.startStep(label);
  try {
    const result = await withRetry(task, {
      ...retryPolicy,
      label,
      onRetry: ({ nextAttempt, attempts, delayMs, error }) => {
        const message = error?.message || "unknown error";
        console.log(`↻ ${label} failed (${message}). retry ${nextAttempt}/${attempts} in ${delayMs}ms`);
        logger?.event("step_retry", {
          step: label,
          nextAttempt,
          attempts,
          delayMs,
          errorType: classifyError(error),
          errorMessage: message,
        });
      },
    });
    logger?.endStep(stepToken, { status: "ok" });
    return result;
  } catch (error) {
    logger?.endStep(stepToken, {
      status: "error",
      errorType: classifyError(error),
      errorMessage: error?.message || String(error),
    });
    throw error;
  }
}

function resolveBasicAuth(config) {
  const fromSaved = loadBetterLetterBasicAuth();
  if (fromSaved?.username && fromSaved?.password) {
    return {
      username: String(fromSaved.username).trim(),
      password: String(fromSaved.password),
      source: "env_or_saved_auth_file",
    };
  }

  const cfgUser = String(config?.betterLetter?.basicAuth?.username || "").trim();
  const cfgPass = String(config?.betterLetter?.basicAuth?.password || "");
  if (cfgUser && cfgPass) {
    return {
      username: cfgUser,
      password: cfgPass,
      source: "runtime_config",
    };
  }

  return null;
}

async function promptForBasicAuth({ logger } = {}) {
  console.log("⚠ BetterLetter Basic Auth not found in environment, saved auth file, or runtime config.");

  const answers = await inquirer.prompt([
    {
      type: "input",
      name: "username",
      message: "Enter BetterLetter Basic Auth username:",
    },
    {
      type: "password",
      name: "password",
      mask: "*",
      message: "Enter BetterLetter Basic Auth password:",
    },
    {
      type: "confirm",
      name: "saveLocal",
      default: true,
      message: "Save these credentials locally for next runs?",
    },
  ]);

  const username = String(answers.username || "").trim();
  const password = String(answers.password || "");

  if (!username || !password) {
    throw new Error(
      "BetterLetter Basic Auth is required. Provide username and password to continue."
    );
  }

  if (answers.saveLocal) {
    const savePath = saveBetterLetterBasicAuth({ username, password });
    console.log(`✔ Saved BetterLetter Basic Auth to ${path.basename(savePath)} (machine-local).`);
    logger?.event("basic_auth_saved", { filePath: savePath });
    return {
      username,
      password,
      source: "interactive_prompt_saved_file",
    };
  }

  console.log(`ℹ Credentials not saved. You can add them later in ${path.basename(BASIC_AUTH_FILE)} or .env.`);
  return {
    username,
    password,
    source: "interactive_prompt",
  };
}

function validateNonInteractiveMode(cliOptions, mode) {
  if (!cliOptions?.nonInteractive) return;

  if (!String(cliOptions.practiceName || "").trim()) {
    throw new Error("Non-interactive mode requires --practice.");
  }

  if (mode === "clean") {
    const cleanType = normalizeCleanType(cliOptions.clean.type);
    const hasSourceFolder = Boolean(String(cliOptions.clean.sourceFolder || "").trim());
    const hasDestinationFolder = Boolean(String(cliOptions.clean.destinationFolder || "").trim());

    if (hasSourceFolder !== hasDestinationFolder) {
      throw new Error(
        "Non-interactive CLEAN manual folder override requires both --source-folder and --destination-folder."
      );
    }

    if (!cleanType && !hasSourceFolder) {
      throw new Error(
        'Non-interactive CLEAN requires --clean-type ("processing" or "filing"), or both --source-folder and --destination-folder.'
      );
    }
    if (!cliOptions.clean.autoConfirm) {
      throw new Error("Non-interactive CLEAN requires --yes (or --confirm-clean).");
    }
  }

  if (mode === "verify") {
    const hasInline = Array.isArray(cliOptions.verify.usernames) && cliOptions.verify.usernames.length > 0;
    const hasFile = Boolean(String(cliOptions.verify.usernamesFile || "").trim());
    if (!hasInline && !hasFile) {
      throw new Error("Non-interactive VERIFY requires --usernames or --usernames-file.");
    }
  }

  if (mode === "create-group") {
    const hasGroupName = Boolean(String(cliOptions.createGroup.groupName || "").trim());
    const hasInline = Array.isArray(cliOptions.verify.usernames) && cliOptions.verify.usernames.length > 0;
    const hasFile = Boolean(String(cliOptions.verify.usernamesFile || "").trim());

    if (!hasGroupName) {
      throw new Error("Non-interactive CREATE GROUP requires --group-name.");
    }
    if (!hasInline && !hasFile) {
      throw new Error("Non-interactive CREATE GROUP requires --usernames or --usernames-file.");
    }
  }
}

async function resolveVerifyUsernames(cliOptions = {}, options = {}) {
  const fromInline = Array.isArray(cliOptions?.verify?.usernames)
    ? cliOptions.verify.usernames
    : [];

  let values = [...fromInline];

  const usernamesFile = String(cliOptions?.verify?.usernamesFile || "").trim();
  if (usernamesFile) {
    if (!fs.existsSync(usernamesFile)) {
      throw new Error(`Usernames file not found: ${usernamesFile}`);
    }
    const raw = fs.readFileSync(usernamesFile, "utf8");
    values = values.concat(splitUsernames(raw));
  }

  values = uniqueUsernames(values);
  if (values.length > 0) return values;

  if (cliOptions?.nonInteractive) {
    throw new Error("No usernames supplied for VERIFY in non-interactive mode.");
  }

  const promptMessage = String(options.promptMessage || "Paste Docman usernames to verify (one per line):");

  const { usernamesRaw } = await inquirer.prompt([
    {
      type: "editor",
      name: "usernamesRaw",
      message: promptMessage,
    },
  ]);

  return uniqueUsernames(splitUsernames(usernamesRaw || ""));
}

function splitUsernames(raw) {
  return String(raw || "")
    .split(/[\n,;]+/)
    .map((value) => value.trim())
    .filter(Boolean);
}

function uniqueUsernames(values = []) {
  const seen = new Set();
  const out = [];

  for (const value of values) {
    const username = String(value || "").trim();
    if (!username) continue;
    const key = username.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(username);
  }

  return out;
}

function isSessionClosed(session, page = session?.page) {
  try {
    if (typeof session?.isClosed === "function" && session.isClosed()) {
      return true;
    }
  } catch (_) {}

  try {
    if (typeof page?.isClosed === "function" && page.isClosed()) {
      return true;
    }
  } catch (_) {}

  return Boolean(session?.closed);
}

async function waitForManualSessionEnd(
  session,
  sessionMessage = "Session finished. Browser will stay open."
) {
  if (isSessionClosed(session)) {
    process.stdin.pause();
    return;
  }

  const context = session?.context;
  const isExternalBrowser = Boolean(session?.isExternalBrowser);
  console.log(`\n${sessionMessage}`);
  if (isExternalBrowser) {
    console.log("Press ENTER here to disconnect automation and exit (Chrome stays open).");
  } else {
    console.log("Close the browser window, or press ENTER here to close it and exit.");
  }

  await new Promise((resolve) => {
    process.stdin.resume();
    process.stdin.once("data", () => resolve());
  });

  await closeSession({ context, safeClose: session?.safeClose });

  process.stdin.pause();
}

async function closeSession(session) {
  if (typeof session?.cleanup === "function") {
    await session.cleanup();
    process.stdin.pause();
    return;
  }

  const safeClose =
    typeof session?.safeClose === "function"
      ? session.safeClose
      : async () => {
          try {
            await session?.context?.close();
          } catch (_) {}
        };

  await safeClose();
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

async function executeMode({
  mode,
  page,
  session,
  cliOptions,
  config,
  runLogger,
}) {
  if (mode === "login") {
    console.log("✅ LOGIN workflow finished.");
    return { outcome: "success" };
  }

  if (mode === "clean") {
    if (cliOptions.clean.autoConfirm) {
      console.log("✔ CLEAN move confirmation will be auto-accepted (--yes)");
    }

    const cleanTypePreset = normalizeCleanType(cliOptions.clean.type);
    const hasManualFolderOverride =
      Boolean(String(cliOptions.clean.sourceFolder || "").trim()) ||
      Boolean(String(cliOptions.clean.destinationFolder || "").trim());

    let cleanType = cleanTypePreset;
    if (!cleanType && !hasManualFolderOverride) {
      cleanType = await promptForCleanType();
    }

    if (!cleanType && cliOptions.clean.type) {
      throw new Error(
        `Unknown CLEAN type "${cliOptions.clean.type}". Use "processing" or "filing".`
      );
    }

    if (cleanType) {
      console.log(`✔ CLEAN type: ${cleanType}`);
    } else if (hasManualFolderOverride) {
      console.log("✔ CLEAN type: manual folder override");
    }

    const cleanSourceFolder = String(
      cliOptions.clean.sourceFolder || config.clean.defaultSourceFolder || ""
    ).trim();
    const cleanDestinationFolder = String(
      cliOptions.clean.destinationFolder || config.clean.defaultDestinationFolder || ""
    ).trim();

    if (cliOptions.clean.sourceFolder) {
      console.log(`✔ CLEAN source folder preset: ${cleanSourceFolder}`);
    }
    if (cliOptions.clean.destinationFolder) {
      console.log(`✔ CLEAN destination folder preset: ${cleanDestinationFolder}`);
    }

    // Each practice names its Docman folders however it likes (e.g. Grove
    // Medical Centre uses "1.For BetterLetter" / "2.Processing by
    // BetterLetter", not either of the two generic conventions this tool
    // otherwise guesses at) - the practice's own EHR Settings page is the
    // actual source of truth, already read once during login (see
    // fetchDocmanCreds.js), so prefer that over guessing. A CLI override
    // still wins if one was given; if neither is available, cleanBetter-
    // LetterProcessing falls back to its built-in guesses.
    const ehrSourceFolder = cleanType === "processing"
      ? String(session?.processingFolder || "").trim()
      : cleanType === "filing"
        ? String(session?.filingFolder || "").trim()
        : "";
    const ehrDestinationFolder = String(session?.inputFolder || "").trim();
    const resolvedSourceFolder = String(cliOptions.clean.sourceFolder || "").trim() || ehrSourceFolder;
    const resolvedDestinationFolder = String(cliOptions.clean.destinationFolder || "").trim() || ehrDestinationFolder;

    if (!cliOptions.clean.sourceFolder && ehrSourceFolder) {
      console.log(`✔ CLEAN source folder from practice's EHR Settings: ${ehrSourceFolder}`);
    }
    if (!cliOptions.clean.destinationFolder && ehrDestinationFolder) {
      console.log(`✔ CLEAN destination folder from practice's EHR Settings: ${ehrDestinationFolder}`);
    }

    if (typeof bootstrapDocmanSession.gotoDocmanFilingAndActivate === "function") {
      console.log("➡ Preparing Docman Filing for CLEAN workflow…");
      await runStepWithRetry({
        label: "prepare_filing_for_clean",
        logger: runLogger,
        retryPolicy: config.retries.step,
        task: async () =>
          bootstrapDocmanSession.gotoDocmanFilingAndActivate(page, {
            skipDialogCheck: true,
            timeouts: config.timeouts,
          }),
      });
    }

    console.log("🧹 Starting CLEAN workflow…");
    await runStepWithRetry({
      label: "clean_workflow",
      logger: runLogger,
      retryPolicy: { attempts: 1, baseDelayMs: 0, maxDelayMs: 0 },
      task: async () =>
        cleanBetterLetterProcessing({
          page,
          cleanType,
          batchSize: config.clean.batchSize,
          dryRun: false,
          defaults: {
            sourceFolder: cleanSourceFolder,
            destinationFolder: cleanDestinationFolder,
          },
          inputs: {
            sourceFolder: resolvedSourceFolder,
            destinationFolder: resolvedDestinationFolder,
            autoConfirmMove: cliOptions.clean.autoConfirm,
            nonInteractive: cliOptions.nonInteractive,
          },
          folderPicker: config.clean.folderPicker,
          retryPolicy: config.retries.step,
          logger: runLogger,
        }),
    });
    console.log("✅ CLEAN workflow finished.");
    return { outcome: "success" };
  }

  if (mode === "verify") {
    console.log("🔍 Starting VERIFY workflow…");

    const usernames = await resolveVerifyUsernames(cliOptions);

    if (!usernames.length) {
      console.log("No usernames provided.");
      return { outcome: "cancelled" };
    }

    console.log(`✔ Verifying ${usernames.length} username(s)`);

    const results = await runStepWithRetry({
      label: "verify_docman_users",
      logger: runLogger,
      retryPolicy: config.retries.step,
      task: async () => verifyDocmanUsers({ page, usernames }),
    });

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
    return { outcome: "success" };
  }

  if (mode === "create-group") {
    console.log("👥 Starting CREATE GROUP workflow…");

    let groupName = String(cliOptions.createGroup.groupName || "").trim();
    if (groupName) {
      console.log(`✔ Group name: ${groupName}`);
    } else {
      const answer = await inquirer.prompt([
        {
          type: "input",
          name: "groupName",
          message: "Enter Docman user group name:",
        },
      ]);
      groupName = String(answer.groupName || "").trim();
    }

    if (!groupName) {
      console.log("No group name provided.");
      return { outcome: "cancelled" };
    }

    const usernames = await resolveVerifyUsernames(cliOptions, {
      promptMessage: "Paste Docman users to add to the group (one per line):",
    });

    if (!usernames.length) {
      console.log("No group members provided.");
      return { outcome: "cancelled" };
    }

    console.log(`✔ Creating "${groupName}" with ${usernames.length} member(s)`);

    const result = await runStepWithRetry({
      label: "create_docman_user_group",
      logger: runLogger,
      retryPolicy: config.retries.step,
      task: async () =>
        createDocmanUserGroup({
          page,
          groupName,
          usernames,
          timeouts: config.timeouts,
          logger: runLogger,
        }),
    });

    console.log(`✅ User group created: ${result.groupName}`);
    if (Array.isArray(result.members) && result.members.length) {
      console.log("Members added:");
      result.members.forEach((member) => console.log(` - ${member}`));
    }
    console.log("✅ CREATE GROUP workflow finished.");
    return { outcome: "success" };
  }

  if (mode === "onboarding") {
    console.log("🧩 Starting ONBOARDING workflow…");
    console.log("ℹ If Docman forced a password change, note the updated password now.");

    let inputFolderName = "zz BL Input. Do not touch";
    if (!cliOptions.nonInteractive) {
      const answer = await inquirer.prompt([
        {
          type: "list",
          name: "inputFolderName",
          message: "Folder #4 choice:",
          default: inputFolderName,
          choices: [
            {
              name: "zz BL Input. Do not touch (default)",
              value: "zz BL Input. Do not touch",
            },
            {
              name: "Not in a folder (practice-specific)",
              value: "Not in a folder",
            },
          ],
        },
      ]);
      inputFolderName = String(answer.inputFolderName || inputFolderName).trim();
    }

    const onboardingResult = await runStepWithRetry({
      label: "onboarding_docman_folders",
      logger: runLogger,
      retryPolicy: config.retries.step,
      task: async () =>
        onboardingDocmanFolders({
          page,
          inputFolderName,
          timeouts: config.timeouts,
          logger: runLogger,
        }),
    });

    if (
      onboardingResult &&
      Number(onboardingResult.existing) === Number(onboardingResult.folderNames?.length || 0)
    ) {
      console.log("ℹ All onboarding folders already exist for this practice. No changes made.");
    }

    console.log("✅ ONBOARDING workflow finished.");
    return { outcome: "success" };
  }

  throw new Error(`Unknown mode selected: ${mode}`);
}

async function promptForModeSelection(config) {
  const answer = await inquirer.prompt([
    {
      type: "list",
      name: "mode",
      message: "What would you like to do next?",
      default: config.modeDefaults.defaultMode,
      choices: getModeChoices(),
    },
  ]);
  return normalizeMode(answer.mode);
}

async function promptForAnotherTask(practiceName) {
  const answer = await inquirer.prompt([
    {
      type: "confirm",
      name: "runAnotherTask",
      default: true,
      message: `Run another task for ${practiceName} in this session?`,
    },
  ]);
  return Boolean(answer.runAnotherTask);
}

function getModeChoices() {
  return [
    { name: "Just login to Docman (no task)", value: "login" },
    { name: "Verify Docman users (copy valid names)", value: "verify" },
    { name: "Create Docman user group", value: "create-group" },
    { name: "Clean BetterLetter folders (processing / filing)", value: "clean" },
    { name: "Onboarding (create BetterLetter filing folders)", value: "onboarding" },
  ];
}

function buildFollowUpCliOptions(cliOptions) {
  return {
    ...cliOptions,
    mode: "",
    clean: {
      type: "",
      sourceFolder: "",
      destinationFolder: "",
      autoConfirm: false,
    },
    verify: {
      usernames: [],
      usernamesFile: "",
    },
    createGroup: {
      groupName: "",
    },
  };
}

function normalizeMode(modeInput) {
  const mode = String(modeInput || "").trim().toLowerCase();
  if (mode === "login" || mode === "login_only") return "login";
  if (mode === "clean") return "clean";
  if (mode === "verify" || mode === "user" || mode === "users") return "verify";
  if (
    mode === "create-group" ||
    mode === "create group" ||
    mode === "create_group" ||
    mode === "creategroup" ||
    mode === "group" ||
    mode === "user-group" ||
    mode === "user group" ||
    mode === "user_group"
  ) {
    return "create-group";
  }
  if (mode === "onboarding") return "onboarding";
  return null;
}

function normalizeCleanType(cleanTypeInput) {
  const cleanType = String(cleanTypeInput || "").trim().toLowerCase();
  if (
    cleanType === "processing" ||
    cleanType === "process" ||
    cleanType === "proc" ||
    cleanType === "processing-folder"
  ) {
    return "processing";
  }
  if (
    cleanType === "filing" ||
    cleanType === "file" ||
    cleanType === "fiiling" ||
    cleanType === "filing-folder"
  ) {
    return "filing";
  }
  return null;
}

async function promptForCleanType() {
  const answer = await inquirer.prompt([
    {
      type: "input",
      name: "cleanType",
      default: "processing",
      message: 'Which CLEAN type? Type "processing" or "filing":',
      validate: (value) =>
        normalizeCleanType(value)
          ? true
          : 'Type "processing" or "filing".',
    },
  ]);

  const cleanType = normalizeCleanType(answer.cleanType);
  if (!cleanType) {
    throw new Error('Unknown CLEAN type. Type "processing" or "filing".');
  }

  return cleanType;
}
