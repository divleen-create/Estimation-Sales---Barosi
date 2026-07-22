@echo off
REM Publish output\index.html to the GitHub Pages repo so the shared URL updates:
REM   https://divleen-create.github.io/Estimation-Sales---Barosi/
setlocal
set "REPO=%~dp0..\site-repo"
if not exist "%REPO%\.git" (
  echo [publish] repo not found at "%REPO%" - clone it first:
  echo   git clone https://github.com/divleen-create/Estimation-Sales---Barosi.git "%REPO%"
  exit /b 1
)
copy /Y "%~dp0output\index.html" "%REPO%\index.html" >nul
pushd "%REPO%"
git add index.html
git diff --cached --quiet && ( echo [publish] no changes. & popd & exit /b 0 )
git -c user.name="divleen-create" -c user.email="divleen-create@users.noreply.github.com" ^
    commit -m "Update index.html (auto refresh)"
git push origin main || ( echo [publish] push failed - check GitHub auth. & popd & exit /b 1 )
popd
echo [publish] live at https://divleen-create.github.io/Estimation-Sales---Barosi/  (updates in ~1 min)
