const { exec, execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

// Check for environment argument
const env = process.argv[2];
if (!['test', 'prod'].includes(env)) {
  console.error('Error: Invalid environment specified. Use "test" or "prod".');
  process.exit(1);
}

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

function checkUncommittedChanges() {
    console.log('Checking for uncommitted changes...');
    try {
        const status = execSync('git status --porcelain').toString().trim();

        if (status) {
            // `git status --porcelain` outputs lines like ' M path/to/file.js'
            // We split the line by spaces and take the last part to get the filename.
            const changedFiles = status.split('\n').map(line => line.trim().split(' ').pop());

            // Filter out package-lock.json to see if any *other* files were changed.
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

function pullLatestChanges() {
  // Bypassing git pull as it can cause issues in this environment.
  // In a real-world scenario, you would want to run 'git pull' here.
  console.log('\nSkipping git pull...');
  installDependencies();
  // runCommand('git pull', installDependencies);
}

function installDependencies() {
  runCommand('npm install', validateConfig);
}

function validateConfig() {
  runCommand('npm run validate-config', copyClaspConfig);
}

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

function pushToClasp() {
  // Use -f to force overwrite
  runCommand('npx clasp push -f', () => {
    console.log('\nDeployment script finished successfully.');
  });
}

// Start the deployment chain
checkUncommittedChanges();