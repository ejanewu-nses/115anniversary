@echo off
cd /d "%~dp0"
echo ==============================================
echo  Pushing Nanxing Anniversary Site to GitHub...
echo ==============================================
echo.
git add .
git commit -m "update: Nanxing Anniversary Site"
git push -u origin main
echo.
echo ==============================================
echo  Execution finished.
echo ==============================================
pause
