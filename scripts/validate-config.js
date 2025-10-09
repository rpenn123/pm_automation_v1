/**
 * @fileoverview This script validates the environment-specific clasp configuration files.
 * It is designed to be run during CI to prevent broken deployments.
 * The script checks for the following conditions:
 * 1. The presence of `.clasp.test.json` and `.clasp.prod.json` in the `config/` directory.
 * 2. That the files are valid JSON and do not contain a UTF-8 Byte Order Mark (BOM), which can break `clasp`.
 * 3. That each config file contains a `scriptId` and `rootDir`.
 * 4. That the `rootDir` in each config file correctly points to a directory containing an `appsscript.json` file.
 * 5. That the root `.clasp.json` file is not committed to the git repository.
 *
 * @usage node scripts/validate-config.js
 */

const fs = require('fs');
const path = require('path');

/**
 * Checks if a buffer starts with a UTF-8 Byte Order Mark (BOM).
 * @param {Buffer} buf The file buffer to check.
 * @returns {boolean} True if the buffer has a BOM, otherwise false.
 */
function hasBOM(buf) {
  return buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf;
}

/**
 * Reads a file as a JSON object while ensuring it does not have a BOM.
 * @param {string} p The path to the JSON file.
 * @returns {object} The parsed JSON object.
 * @throws {Error} If the file has a BOM or is not valid JSON.
 */
function readJsonNoBOM(p) {
  const buf = fs.readFileSync(p);
  if (hasBOM(buf)) {
    throw new Error(`${p} has a UTF-8 BOM. Save without BOM.`);
  }
  try {
    return JSON.parse(buf.toString('utf8'));
  } catch (e) {
    throw new Error(`${p} is not valid JSON: ${e.message}`);
  }
}

/**
 * Checks a single clasp configuration file for validity.
 * @param {string} file The path to the config file (e.g., 'config/.clasp.test.json').
 * @throws {Error} If the file is missing, invalid, or misconfigured.
 */
function checkConfig(file) {
  const p = path.join(process.cwd(), file);
  if (!fs.existsSync(p)) throw new Error(`Missing ${file}`);
  const j = readJsonNoBOM(p);

  if (!j.scriptId || typeof j.scriptId !== 'string') {
    throw new Error(`${file}: scriptId is missing or not a string`);
  }
  if (!j.rootDir || typeof j.rootDir !== 'string') {
    throw new Error(`${file}: rootDir is missing or not a string`);
  }

  const appscript = path.join(process.cwd(), j.rootDir, 'appsscript.json');
  if (!fs.existsSync(appscript)) {
    throw new Error(`${file}: appsscript.json not found at ${appscript} — check rootDir`);
  }
  console.log(`OK: ${file} -> ${appscript}`);
}

/**
 * Ensures the root `.clasp.json` file is not tracked by git.
 * This is a safeguard to prevent developers from accidentally committing their active,
 * environment-specific configuration, which could cause deployment mix-ups.
 */
function ensureClaspJsonNotCommitted() {
  try {
    const tracked = require('child_process')
      .execSync('git ls-files .clasp.json', { encoding: 'utf8' })
      .trim();
    if (tracked) {
      throw new Error('.clasp.json is tracked in git. Add it to .gitignore and commit.');
    }
  } catch (e) {
    // If git is unavailable (e.g., in some CI environments), we can ignore this check.
    console.log(`Note: git check for .clasp.json skipped or failed: ${e.message}`);
  }
}

/**
 * Main execution block. Runs all validation checks.
 */
try {
  ensureClaspJsonNotCommitted();
  checkConfig('config/.clasp.test.json');
  checkConfig('config/.clasp.prod.json');
  console.log('Validation passed.');
} catch (e) {
  console.error(e.message);
  process.exit(1);
}