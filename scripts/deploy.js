// scripts/deploy.js
const fs = require('fs-extra');
const path = require('path');
const { exec } = require('child_process');

function deploy() {
  const env = process.argv[2];
  if (!['test', 'prod'].includes(env)) {
    console.error('Usage: node scripts/deploy.js [test|prod]');
    process.exit(1);
  }

  console.log('Pulling latest changes from Git...');
  const gitPull = exec('git pull');

  gitPull.stdout.pipe(process.stdout);
  gitPull.stderr.pipe(process.stderr);

  gitPull.on('close', (code) => {
    if (code !== 0) {
      console.error(`\nError: git pull exited with code ${code}`);
      process.exit(1);
    }

    console.log('\nGit pull successful. Starting deployment...');
    runDeployment(env);
  });
}

async function runDeployment(env) {
  const rootDir = process.cwd();
  const configDir = path.join(rootDir, 'config');
  const claspConfigFile = `.clasp.${env}.json`;
  const sourcePath = path.join(configDir, claspConfigFile);
  const destPath = path.join(rootDir, '.clasp.json');

  try {
    // 1. Check if the source config file exists
    if (!await fs.pathExists(sourcePath)) {
      console.error(`Error: Config file not found at ${sourcePath}`);
      process.exit(1);
    }

    // 2. Copy the correct config file to the root
    console.log(`Copying ${sourcePath} to ${destPath}...`);
    await fs.copy(sourcePath, destPath, { overwrite: true });

    // 3. Run clasp push
    console.log(`Deploying to ${env.toUpperCase()} environment...`);
    const claspCommand = 'npx clasp push -f'; // -f to force overwrite
    const child = exec(claspCommand);

    // 4. Stream output to console
    child.stdout.pipe(process.stdout);
    child.stderr.pipe(process.stderr);

    child.on('close', (code) => {
      if (code === 0) {
        console.log(`\nDeployment to ${env.toUpperCase()} completed successfully.`);
      } else {
        console.error(`\nError: clasp push exited with code ${code}`);
        process.exit(1);
      }
    });
  } catch (error) {
    console.error('An error occurred during deployment:', error);
    process.exit(1);
  }
}

deploy();