/**
 * @fileoverview This script automates the deployment of the Google Apps Script project.
 * It ensures a safe and consistent deployment by performing a series of checks and operations in sequence:
 * 1. Verifies that the correct environment ('test' or 'prod') is specified.
 * 2. Checks for any uncommitted git changes (excluding package-lock.json) and aborts if any are found.
 * 3. Pulls the latest changes from the git repository.
 * 4. Installs npm dependencies.
 * 5. Validates the clasp configuration files using `validate-config.js`.
 * 6. Copies the appropriate environment-specific `.clasp.[env].json` file to `.clasp.json`.
 * 7. Pushes the code to the corresponding Google Apps Script project using `clasp push`.
 *
 * @usage node scripts/deploy.js <test|prod>
 */

const { exec, execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

// Check for environment argument
const env = process.argv[2];
if (!['test', 'prod'].includes(env)) {
  console.error('Error: Invalid environment specified. Use "test" or "prod".');
  process.exit(1);
}

/**
 * Executes a shell command asynchronously and pipes its output to the console.
 * Exits the process if the command fails.
 * @param {string} command The shell command to execute.
 * @param {function} [onSuccess] A callback function to execute if the command succeeds.
 */
function runCommand(command, onSuccess) {
  console.log(`\n> ${command}`);
  const childProcess = exec(command);

  // Pipe stdout and stderr to the parent process to see the output in real-time
  childProcess.stdout.pipe(process.stdout);
  childProcess.stderr.pipe(process.stderr);

  childProcess.on('close', (code) => {
    if (code !== 0) {
      console.error(`\nError: Command "${command}" exited with code ${code}`);
      process.exit(1);
    } else if (onSuccess) {
      onSuccess();
    }
  });
}

/**
 * Checks for uncommitted changes in the git repository, ignoring package-lock.json.
 * If changes are found, the script aborts to prevent an unsafe deployment.
 * If the working directory is clean, it proceeds to the next step in the deployment chain.
 */
function checkUncommittedChanges() {
    console.log('Checking for uncommitted changes...');
    try {
        const status = execSync('git status --porcelain').toString().trim();
        if (status) {
            const changedFiles = status.split('\n').map(line => line.trim().split(' ').pop());
            const otherChanges = changedFiles.filter(file => file !== 'package-lock.json');
            if (otherChanges.length > 0) {
                console.error('Error: You have uncommitted changes. Please commit or stash them before deploying.');
                console.error('Uncommitted files:', otherChanges.join(', '));
                process.exit(1);
            }
        }
        console.log('No uncommitted changes found (ignoring package-lock.json).');
        pullLatestChanges();
    } catch (error) {
        console.error('Error checking for uncommitted changes:', error);
        process.exit(1);
    }
}

/**
 * Pulls the latest changes from the current git branch.
 */
function pullLatestChanges() {
  console.log('\nPulling latest changes from git...');
  runCommand('git pull', installDependencies);
}

/**
 * Installs or updates npm dependencies.
 */
function installDependencies() {
  runCommand('npm install', validateConfig);
}

/**
 * Runs the configuration validation script.
 */
function validateConfig() {
  runCommand('npm run validate-config', copyClaspConfig);
}

/**
 * Copies the environment-specific clasp configuration file to the root `.clasp.json`.
 * This is the mechanism that directs `clasp` to the correct Google Apps Script project.
 */
function copyClaspConfig() {
  console.log('\nCopying clasp config file...');
  const rootDir = process.cwd();
  const configDir = path.join(rootDir, 'config');
  const claspConfigFile = `.clasp.${env}.json`;
  const sourcePath = path.join(configDir, claspConfigFile);
  const destPath = path.join(rootDir, '.clasp.json');

  if (!fs.existsSync(sourcePath)) {
    console.error(`Error: Clasp config file not found at ${sourcePath}`);
    process.exit(1);
  }
  fs.copyFileSync(sourcePath, destPath);
  console.log(`Copied ${claspConfigFile} to .clasp.json successfully.`);
  pushToClasp();
}

/**
 * Pushes the source code from the `src` directory to the linked Google Apps Script project.
 * Uses the `-f` flag to force an overwrite of the remote project.
 */
function pushToClasp() {
  runCommand('npx clasp push -f', () => {
    console.log('\nDeployment script finished successfully.');
  });
}

// Start the deployment chain
checkUncommittedChanges();