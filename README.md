# Google Sheets Automation & Dashboard

[![Validate Deploy Configs](https://github.com/rpenn123/pm_automation_v1/actions/workflows/validate-deploy.yml/badge.svg)](https://github.com/rpenn123/pm_automation_v1/actions/workflows/validate-deploy.yml)

This repository contains the Google Apps Script (GAS) codebase for a powerful project management automation and dashboard system. It is designed to be synced with Google Sheets and includes a streamlined, cross-platform deployment process.

## 1. Quick Start

This guide provides the essential commands to get the project running locally.

### Prerequisites

*   **Node.js:** The latest Long-Term Support (LTS) version is recommended.
*   **Git:** For version control.
*   **`clasp` Login:** Authenticate `clasp` with your Google account by running: `npx clasp login`. You only need to do this once.
*   **Google Apps Script API:** Ensure the API is enabled for your Google account. You can do so [here](https://script.google.com/home/usersettings).

### Installation & Deployment

1.  **Clone the repository:**
    ```sh
    git clone https://github.com/rpenn123/pm_automation_v1.git
    cd pm_automation_v1
    ```

2.  **Install dependencies:**
    ```sh
    npm install
    ```

3.  **Deploy to an environment:**
    *   **To TEST:** `npm run deploy:test`
    *   **To PRODUCTION:** `npm run deploy:prod`

    These commands copy the correct `.clasp.[env].json` configuration from the `config/` directory and push the `src/` folder to the corresponding Apps Script project.

## 2. Project Structure

The repository is organized to separate concerns, making it easier to navigate and maintain.

```
.
├── .github/          # CI workflows
├── config/           # Environment-specific .clasp.json files
├── docs/             # Project documentation, including the pre-audit README
├── scripts/          # Node.js scripts for deployment and validation
├── src/              # Google Apps Script source code
│   ├── core/         # Core logic (onEdit router, transfer engine)
│   ├── services/     # External services (logging, last edit tracking)
│   └── ui/           # UI-related code (dashboard, custom menus)
├── tests/            # Manual testing plans
├── .gitignore        # Specifies intentionally untracked files
├── package.json      # Defines scripts and dependencies
└── README.md         # This file
```

## 3. Test Matrix

Currently, testing is performed manually. The smoke test plan provides a consistent way to verify core functionality after a deployment.

*   **Smoke Test:** A detailed plan is located in `tests/smoke-test.md`. This should be run against the TEST environment after every deployment.

## 4. CI/Deployment

*   **CI:** A GitHub Actions workflow runs on every pull request to `main`. It executes `npm run validate-config` to ensure that the `clasp` configuration files in `config/` are valid and correctly formatted.
*   **Deployment:** Deployments are handled via the `npm run deploy:*` commands as described in the Quick Start. This process is designed to be run from a developer's local machine.

## 5. Contribution Guide

1.  Create a new feature branch from `main`.
2.  Make your code changes within the `src/` directory.
3.  If you modify deployment or validation logic, update the scripts in `scripts/`.
4.  Run `npm run deploy:test` to deploy your changes to the TEST environment.
5.  Perform the smoke test in `tests/smoke-test.md` to verify functionality.
6.  Open a pull request to `main`.

## 6. Rollback and Revert Procedure

If a deployment introduces a critical issue, the safest way to roll back is to revert the pull request that contained the breaking change.

1.  In GitHub, find and revert the PR.
2.  Check out the `main` branch locally and pull the revert commit.
3.  Deploy the stable version to the affected environment: `npm run deploy:prod`.

---
*Owner: Ryan (rpenn@mobility123.com)*