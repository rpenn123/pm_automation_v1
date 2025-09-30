#!/bin/bash
# Idempotent script to refactor the project structure.
# Exit immediately if a command exits with a non-zero status.
set -e

# 1. Create new directories if they don't exist
mkdir -p src/core src/services src/ui tests config docs

# 2. Move files to their new locations, checking for existence first
[ -f "src/Automations.gs" ] && git mv src/Automations.gs src/core/
[ -f "src/TransferEngine.gs" ] && git mv src/TransferEngine.gs src/core/
[ -f "src/Utilities.gs" ] && git mv src/Utilities.gs src/core/
[ -f "src/LastEditService.gs" ] && git mv src/LastEditService.gs src/services/
[ -f "src/LoggerService.gs" ] && git mv src/LoggerService.gs src/services/
[ -f "src/Dashboard.gs" ] && git mv src/Dashboard.gs src/ui/
[ -f "src/Setup.gs" ] && git mv src/Setup.gs src/ui/
[ -f "TESTING_PLAN.md" ] && git mv TESTING_PLAN.md tests/smoke-test.md
[ -f "scripts/validate-deploy.js" ] && git mv scripts/validate-deploy.js scripts/validate-config.js
[ -f "README.md" ] && git mv README.md docs/README.pre-audit.md

# 3. Handle config files
# These are not committed, so we just move them.
# The deploy script will handle the active .clasp.json
if [ -f ".clasp.test.json" ]; then
    mv .clasp.test.json config/
fi
if [ -f ".clasp.prod.json" ]; then
    mv .clasp.prod.json config/
fi

# 4. Remove old, now-empty directories and scripts
rm -f update.ps1 update_prod.ps1 update_production.bat update_test.bat src/AllTests.gs

echo "Structural refactoring complete. The following steps must be done manually:"
echo "1. Create a new package.json and install dependencies."
echo "2. Create the new deploy.js script."
echo "3. Create the top-level README.md."
echo "4. Update .github/workflows/validate-deploy.yml to use the new script path."
echo "5. Create a .gitignore file."