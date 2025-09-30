// scripts/validate-config.js
// Fails CI if config/.clasp.*.json is invalid, has a BOM, or points to a missing appsscript.json

const fs = require('fs');
const path = require('path');

function hasBOM(buf) {
  return buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf;
}

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

function ensureClaspJsonNotCommitted() {
  // The active .clasp.json should not be tracked to avoid env confusion
  try {
    const tracked = require('child_process')
      .execSync('git ls-files .clasp.json', { encoding: 'utf8' })
      .trim();
    if (tracked) {
      throw new Error('.clasp.json is tracked in git. Add it to .gitignore and commit.');
    }
  } catch (e) {
    // If git is unavailable in CI, ignore this check
    console.log(`Note: git check skipped or failed: ${e.message}`);
  }
}

try {
  ensureClaspJsonNotCommitted();
  checkConfig('config/.clasp.test.json');
  checkConfig('config/.clasp.prod.json');
  console.log('Validation passed.');
} catch (e) {
  console.error(e.message);
  process.exit(1);
}