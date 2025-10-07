@echo off
echo =================================================
echo             DEPLOYING TO TEST
echo =================================================
echo.

node scripts/deploy.js test

if %errorlevel% neq 0 (
  echo.
  echo ERROR: Deployment to TEST failed.
  pause
  exit /b 1
)

echo.
echo =================================================
echo           DEPLOYMENT SUCCESSFUL
echo =================================================
echo.
pause