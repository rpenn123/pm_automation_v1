// scripts/deploy.js
const fs = require('fs-extra');
const path = require('path');
const { exec } = require('child_process');
const util = require('util');

const execPromise = util.promisify(exec);

async function runCommand(command, logMessage) {
  console.log(logMessage);
  try {
    const { stdout, stderr } = await execPromise(command);
    if (stdout) console.log(stdout);
    if (stderr) console.error(stderr);
    console.log(`${logMessage} - Success`);
  } catch (error) {
    console.error(`Error during: "${logMessage}"`);
    console.error(error.message);
    throw new Error(`Failed to execute command: ${command}`);
  }
}

async function checkUncommittedChanges() {
  console.log('Checking for uncommitted changes...');
  const { stdout } = await execPromise('git status --porcelain');
  if (stdout) {
    console.error('ERROR: You have uncommitted changes.');
    console.error('Please commit or stash them before deploying.');
    console.error(stdout);
    throw new Error('Uncommitted changes found.');
  }
  console.log('No uncommitted changes found.');
}

async function deploy() {
  const env = process.argv[2];
  if (!['test', 'prod'].includes(env)) {
    console.error('Usage: node scripts/deploy.js [test|prod]');
    process.exit(1);
  }

  try {
    // Step 1: Pull latest changes from Git
    await runCommand('git pull', 'Pulling latest changes from Git...');

    // Step 2: Install/update dependencies
    await runCommand('npm install', 'Installing dependencies...');

    // Step 3: Check for uncommitted changes (post-pull and post-install)
    await checkUncommittedChanges();

    // Step 4: Validate clasp configuration
    await runCommand('npm run validate-config', 'Validating clasp configuration...');

    // Step 5: Copy the correct config file to the root
    const rootDir = process.cwd();
    const configDir = path.join(rootDir, 'config');
    const claspConfigFile = `.clasp.${env}.json`;
    const sourcePath = path.join(configDir, claspConfigFile);
    const destPath = path.join(rootDir, '.clasp.json');

    console.log(`Copying ${sourcePath} to ${destPath}...`);
    if (!await fs.pathExists(sourcePath)) {
      throw new Error(`Config file not found at ${sourcePath}`);
    }
    await fs.copy(sourcePath, destPath, { overwrite: true });
    console.log('Config file copied successfully.');

    // Step 6: Run clasp push
    await runCommand('npx clasp push -f', `Deploying to ${env.toUpperCase()} environment...`);

    console.log(`\nDeployment to ${env.toUpperCase()} completed successfully.`);

  } catch (error) {
    console.error('\n--- DEPLOYMENT FAILED ---');
    console.error(error.message);
    process.exit(1);
  }
}

deploy();