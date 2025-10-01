@echo off
echo =================================================
echo  Deploying to PRODUCTION Environment
echo =================================================
echo.
echo Running npm run deploy:prod...
call npm run deploy:prod
echo.
echo =================================================
echo  Deployment to PRODUCTION Complete
echo =================================================
pause