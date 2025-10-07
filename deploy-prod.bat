@echo off
echo =================================================
echo           DEPLOYING TO PRODUCTION
echo =================================================
echo.

node scripts/deploy.js prod

if %errorlevel% neq 0 (
  echo.
  echo ERROR: Deployment to PRODUCTION failed.
  exit /b 1
)

echo.
echo =================================================
echo           DEPLOYMENT SUCCESSFUL
echo =================================================
echo.
pause