const path = require("path");

function parseCliArgs(argv = process.argv.slice(2)) {
  const out = {
    help: false,
    nonInteractive: false,
    practiceName: "",
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

  const args = [...argv];

  for (let i = 0; i < args.length; i++) {
    const token = String(args[i] || "").trim();
    if (!token) continue;

    const { flag, inlineValue } = splitFlagAndValue(token);

    switch (flag) {
      case "-h":
      case "--help":
        out.help = true;
        break;

      case "--non-interactive":
        out.nonInteractive = true;
        break;

      case "-p":
      case "--practice": {
        const value = readValue(args, i, inlineValue, flag);
        if (inlineValue == null) i += 1;
        out.practiceName = value;
        break;
      }

      case "-m":
      case "--mode": {
        const value = readValue(args, i, inlineValue, flag);
        if (inlineValue == null) i += 1;
        out.mode = value;
        break;
      }

      case "--clean-type": {
        const value = readValue(args, i, inlineValue, flag);
        if (inlineValue == null) i += 1;
        out.clean.type = value;
        break;
      }

      case "--source":
      case "--source-folder": {
        const value = readValue(args, i, inlineValue, flag);
        if (inlineValue == null) i += 1;
        out.clean.sourceFolder = value;
        break;
      }

      case "--destination":
      case "--destination-folder": {
        const value = readValue(args, i, inlineValue, flag);
        if (inlineValue == null) i += 1;
        out.clean.destinationFolder = value;
        break;
      }

      case "-y":
      case "--yes":
      case "--confirm-clean":
        out.clean.autoConfirm = true;
        break;

      case "--usernames": {
        const value = readValue(args, i, inlineValue, flag);
        if (inlineValue == null) i += 1;
        out.verify.usernames = uniqueNormalizedUsernames([
          ...out.verify.usernames,
          ...splitUsernames(value),
        ]);
        break;
      }

      case "--usernames-file": {
        const value = readValue(args, i, inlineValue, flag);
        if (inlineValue == null) i += 1;
        out.verify.usernamesFile = value;
        break;
      }

      case "--group-name":
      case "--group": {
        const value = readValue(args, i, inlineValue, flag);
        if (inlineValue == null) i += 1;
        out.createGroup.groupName = value;
        break;
      }

      default:
        throw new Error(`Unknown option: ${flag}`);
    }
  }

  if (out.verify.usernamesFile) {
    out.verify.usernamesFile = path.resolve(process.cwd(), out.verify.usernamesFile);
  }

  return out;
}

function splitFlagAndValue(token) {
  if (!token.startsWith("--")) {
    return { flag: token, inlineValue: null };
  }

  const eqIndex = token.indexOf("=");
  if (eqIndex === -1) return { flag: token, inlineValue: null };

  const flag = token.slice(0, eqIndex);
  const inlineValue = token.slice(eqIndex + 1);
  return { flag, inlineValue };
}

function readValue(args, index, inlineValue, flag) {
  const raw = inlineValue != null ? inlineValue : args[index + 1];
  if (raw == null) {
    throw new Error(`Missing value for ${flag}`);
  }

  const value = String(raw).trim();
  if (!value) {
    throw new Error(`Empty value for ${flag}`);
  }

  if (inlineValue == null && value.startsWith("-")) {
    throw new Error(`Missing value for ${flag}`);
  }

  return value;
}

function splitUsernames(raw) {
  return String(raw || "")
    .split(/[\n,;]+/)
    .map((value) => value.trim())
    .filter(Boolean);
}

function uniqueNormalizedUsernames(values) {
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

function printCliHelp() {
  console.log("Docman Tool CLI options:\n");
  console.log("  --help, -h");
  console.log("      Show this help and exit.\n");
  console.log("  --practice, -p <name>");
  console.log("      Pre-fill practice name (Step 0).\n");
  console.log("  --mode, -m <login|clean|verify|create-group|onboarding>");
  console.log("      Pre-select workflow mode.\n");
  console.log("  --clean-type <processing|filing>");
  console.log("      CLEAN subtype. Auto-resolves BetterLetter folders for this cleanup.\n");
  console.log("  --group-name, --group <name>");
  console.log("      User group name for CREATE GROUP mode.\n");
  console.log("  --source-folder, --source <name>");
  console.log("      Optional manual CLEAN source folder override.\n");
  console.log("  --destination-folder, --destination <name>");
  console.log("      Optional manual CLEAN destination folder override.\n");
  console.log("  --yes, -y, --confirm-clean");
  console.log("      Auto-confirm CLEAN prompts.\n");
  console.log("  --usernames \"u1,u2\"");
  console.log("      User list for VERIFY or CREATE GROUP (comma/newline/semicolon separated).\n");
  console.log("  --usernames-file <path>");
  console.log("      User list from a text file for VERIFY or CREATE GROUP (one per line).\n");
  console.log("  --non-interactive");
  console.log("      Fail fast instead of prompting for missing inputs.\n");
  console.log("Examples:");
  console.log('  node run.js --practice "Ashfield" --mode login');
  console.log('  node run.js -p "Heathview" -m clean --clean-type processing -y');
  console.log('  node run.js -p "Heathview" -m clean --clean-type filing -y');
  console.log('  node run.js -p "Heathview" -m clean --source "BetterLetter: Processing" --destination "BetterLetter: Input" -y');
  console.log('  node run.js -p "Queen Square" -m verify --usernames-file ./inputs/users.txt');
  console.log('  node run.js -p "Ribblesdale Medical Practice" -m create-group --group-name "All Doctors" --usernames-file ./inputs/users.txt');
  console.log('  node run.js -p "Edlesborough" -m onboarding');
}

module.exports = {
  parseCliArgs,
  printCliHelp,
};
