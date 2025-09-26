# Optimized Google Sheets Automation & Dashboard — Ops & Deployment

This repo contains a Google Apps Script (GAS) codebase synced with a bound Script on two Sheets (TEST/PROD). It includes one-click Windows launchers that swap `.clasp.json` and push safely via `clasp`.

## What this system does
- Two-way **Progress** sync between `Forecasting` and `Upcoming`
- Automated row transfers (Permits approved → `Upcoming`; Delivered TRUE → `Inventory_Elevators`; In Progress → `Framing`)
- **Last Edit** tracking, external **monthly audit logs**, error **email alerts**
- Menu-driven **Dashboard** with drilldowns and charts

## Environments
- **TEST Script ID:** `1vYXETLX8I3HICSveg7FmLKWbZiwToKicThNeOxm_maQgnQ97AwDmz7iX`
- **PROD Script ID:** `15_PrYM6MxfCbA1bXt0deEGI7cXf74B_KlJY7Ydw59uVrmHZn3IEKFGPJ`

> The active target is controlled by `.clasp.json`, which our launchers overwrite from `.clasp.test.json` or `.clasp.prod.json`.

## Prereqs
- Node.js LTS and `@google/clasp` (`npm i -g @google/clasp`, then `clasp login`)
- Git
- Editor access to both Sheets
- Apps Script API enabled for your Google account

## Repo layout (current)

> **Note on `rootDir`:** Our `.clasp.*.json` files store the detected code root (that long `Optimized-…/src` path). The deploy scripts auto-detect `appsscript.json` if `rootDir` ever drifts, so pushes don’t break.

## Daily use
- **TEST deploy:** double-click `update_test.bat` (or `.\update_test.bat` in PowerShell)
- **PROD deploy:** double-click `update_production.bat`
- After deploy: open the target Sheet → refresh → **Project Actions → Run Full Setup** if scopes/triggers changed

## Smoke test (2 minutes)
1. Edit `Progress` for a row mirrored in `Upcoming` → verify sync
2. Set `Permits=approved` → appears in `Upcoming`
3. Set `Delivered=TRUE` → appears in `Inventory_Elevators`
4. Set `Progress=In Progress` → appears in `Framing`
5. Check **Extensions → Apps Script → Executions** for green runs

## CI guard (already in repo)
The workflow `.github/workflows/validate-deploy.yml` runs `scripts/validate-deploy.js` to fail any PR/commit that:
- has BOM in `.clasp.*.json`
- has invalid JSON
- points `rootDir` at a folder that does not contain `appsscript.json`
- accidentally tracks `.clasp.json` in git

## Rollback
- Checkout a known-good commit → deploy to TEST → validate → deploy to PROD.

## Recommended hardening (next steps)
1. Enable **branch protection** on `main`: require PR, require the “Validate deploy configs” check, block force-push and branch deletes.
2. (Optional) Flatten repo so code lives at `/src` (simpler `rootDir`), then update `.clasp.*.json` to `"./src"`.
3. Add **CODEOWNERS** (e.g., `*  @rpenn123`) so PRs need your review.
4. Add `docs/RELEASE.md` with the smoke-test checklist above.

— Owner: Ryan (rpenn@mobility123.com)

