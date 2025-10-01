@echo off
echo =================================================
echo  Deploying to TEST Environment
echo =================================================
echo.
echo Running npm run deploy:test...
call npm run deploy:test
echo.
echo =================================================
echo  Deployment to TEST Complete
echo =================================================
pause