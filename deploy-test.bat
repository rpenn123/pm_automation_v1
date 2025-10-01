@echo off
echo =================================================
echo  Deploying to TEST Environment
echo =================================================
echo.
call npm run deploy:test
if %errorlevel% neq 0 (
  echo.
  echo ERROR: Deployment script failed.
  pause
  exit /b %errorlevel%
)
echo.
echo =================================================
echo  Deployment to TEST Complete
echo =================================================
pause