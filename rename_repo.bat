@echo off
setlocal EnableDelayedExpansion
REM ===========================================================================
REM  fix-repo-remote.bat
REM  Repairs the "origin" remote URL of a repo after the GitHub repo was renamed.
REM
REM  Use:  put this .bat in the PARENT folder that contains the repo subfolder,
REM        double-click it (or run it from that folder).
REM
REM  Args (optional):
REM        %1 = repo subfolder name   (default below; auto-detected if omitted)
REM        /y = skip the confirmation prompt
REM
REM  Exit codes: 0 = ok/already correct, 1 = error, 2 = aborted by user
REM ===========================================================================

set "OWNER=VoidForge"
set "NEWNAME=csi_safe_automation"
set "OLDNAME=csi_safe_vba"
set "NEWURL=https://github.com/%OWNER%/%NEWNAME%.git"

set "BASE=%CD%"
set "REPO_DIR=%~1"
set "YES="
if /i "%~1"=="/y" (set "REPO_DIR=" & set "YES=1")
if /i "%~2"=="/y" set "YES=1"
if "%REPO_DIR%"=="" set "REPO_DIR=%NEWNAME%"

set "TARGET=%BASE%\%REPO_DIR%"

echo ============================================================
echo  Fix git remote URL after repo rename
echo ============================================================
echo  Base dir : %BASE%
echo  Repo dir : %REPO_DIR%
echo  New URL  : %NEWURL%
echo.

REM --- locate the repo if the default subfolder name was wrong --------------
if not exist "%TARGET%\.git" (
    set "COUNT=0"
    set "FOUND="
    for /d %%D in ("%BASE%\*") do (
        if exist "%%D\.git" (
            set "FOUND=%%D"
            set /a COUNT+=1
        )
    )
    if !COUNT!==1 (
        echo [INFO] "%REPO_DIR%" not found - using detected repo "!FOUND!"
        set "TARGET=!FOUND!"
    ) else if !COUNT! gtr 1 (
        echo [ERROR] No .git in "%TARGET%" and multiple repos found here.
        echo         Pass the folder name:  %~nx0 ^<repo-folder^>
        goto :fail
    )
)

if not exist "%TARGET%\.git" (
    echo [ERROR] No git repository at "%TARGET%".
    goto :fail
)

where git >nul 2>&1 || (echo [ERROR] git is not on PATH. & goto :fail)

cd /d "%TARGET%" || (echo [ERROR] Cannot enter "%TARGET%". & goto :fail)

REM --- read current origin --------------------------------------------------
set "CUR="
for /f "delims=" %%U in ('git config --get remote.origin.url 2^>nul') do set "CUR=%%U"

if "%CUR%"=="" (
    echo [ERROR] This repo has no "origin" remote.
    goto :fail
)
echo  Current   : %CUR%

if /i "%CUR%"=="%NEWURL%" (
    echo.
    echo [OK] origin already points at the new URL - nothing to do.
    goto :done
)

set "CHK=!CUR:%OLDNAME%=!"
if "!CHK!"=="!CUR!" (
    echo [WARN] Current URL does not contain "%OLDNAME%".
    echo        Only continue if you are sure this is the right repository.
) else (
    echo [INFO] Old repo name detected - updating.
)

if not "%YES%"=="1" (
    choice /c YN /m "Set origin to %NEWURL%"
    if errorlevel 2 (echo [SKIP] Aborted by user. & goto :abort)
)

git remote set-url origin "%NEWURL%"
if errorlevel 1 (echo [ERROR] "git remote set-url" failed. & goto :fail)

echo.
echo Verifying...
git remote -v
git fetch origin
if errorlevel 1 (echo [ERROR] "git fetch" failed - check URL and credentials. & goto :fail)

echo.
echo [OK] origin updated and fetch succeeded.
goto :done

:fail
cd /d "%BASE%"
endlocal & exit /b 1

:abort
cd /d "%BASE%"
endlocal & exit /b 2

:done
cd /d "%BASE%"
endlocal & exit /b 0