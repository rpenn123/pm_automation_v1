@echo off
echo =================================================
echo  Deploying to PRODUCTION Environment
echo =================================================
echo.
call npm run deploy:prod
if %errorlevel% neq 0 (
  echo.
  echo ERROR: Deployment script failed.
  pause
  exit /b %errorlevel%
)
echo.
echo =================================================
echo  Deployment to PRODUCTION Complete
echo =================================================
pause