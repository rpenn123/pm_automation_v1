@echo off
setlocal

echo =================================================
echo  Preparing for PRODUCTION Environment Deployment
echo =================================================
echo.

echo Checking for uncommitted changes...
for /f "tokens=*" %%a in ('git status --porcelain') do (
  echo.
  echo WARNING: You have uncommitted changes.
  echo Please commit or stash them before deploying.
  git status -s
  echo.
  pause
  goto:eof
)
echo No uncommitted changes found.
echo.

echo Installing dependencies...
call npm install
if %errorlevel% neq 0 (
  echo ERROR: npm install failed.
  pause
  exit /b %errorlevel%
)
echo.

echo Validating clasp configuration...
call npm run validate-config
if %errorlevel% neq 0 (
  echo ERROR: Configuration validation failed.
  pause
  exit /b %errorlevel%
)
echo.

echo =================================================
echo  Deploying to PRODUCTION Environment
echo =================================================
echo.
call npm run deploy:prod
if %errorlevel% neq 0 (
  echo ERROR: Deployment to PRODUCTION failed.
  pause
  exit /b %errorlevel%
)
echo.

echo =================================================
echo  Deployment to PRODUCTION Complete
echo =================================================
pause