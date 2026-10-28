@echo off
REM post_build.cmd - XLEdge (shell)
REM
REM Developer loop: stage the core release that XLEdge.Addin.Core's post-build published
REM (XLEdge.Addin.Core\SetupFiles\<Config>\Manifest\) into this project's bin\<Config>\AddinCore\Manifest\,
REM so the next Excel start (or Reload Add-in) installs it. Old zips are removed first: the bootstrapper
REM uses the manifest's fileName, but a single staged zip keeps the folder unambiguous.
REM The quoted set "VAR=value" form keeps paths with spaces, '&' or '(' intact.

set "TARGET_DIR=%~1"
if "%TARGET_DIR:~-1%"=="\" set "TARGET_DIR=%TARGET_DIR:~0,-1%"
set "PROJECT_DIR=%~dp0"
set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
for %%F in ("%PROJECT_DIR%") do set "SOLUTION_DIR=%%~dpF"
set "SOLUTION_DIR=%SOLUTION_DIR:~0,-1%"
for %%F in ("%TARGET_DIR%") do set "CONFIG=%%~nxF"

set "CORE_PACKAGE=%SOLUTION_DIR%\XLEdge.Addin.Core\SetupFiles\%CONFIG%\Manifest"
set "STAGED=%TARGET_DIR%\AddinCore\Manifest"

if not exist "%CORE_PACKAGE%\manifest.json" (
    echo warning XLEDGE0002: no core package in "%CORE_PACKAGE%" - build XLEdge.Addin.Core first.
    exit /b 0
)
if not exist "%STAGED%" mkdir "%STAGED%"
if errorlevel 1 (
    echo error XLEDGE0003: could not create "%STAGED%".
    exit /b 1
)
del /q "%STAGED%\*.zip" 2>nul
xcopy /Y /I /Q "%CORE_PACKAGE%\*" "%STAGED%\" >nul
if errorlevel 1 (
    echo error XLEDGE0003: could not stage the core package into "%STAGED%".
    exit /b 1
)

REM Add-in Express loader (registered by the installer; needed next to XLEdge.dll for a dev registration too)
xcopy /Y /I /Q "%PROJECT_DIR%\Loader\*" "%TARGET_DIR%\" >nul
if errorlevel 1 (
    echo error XLEDGE0004: could not copy the Add-in Express loader into "%TARGET_DIR%".
    exit /b 1
)

echo XLEdge post-build: staged the core package into %STAGED%
exit /b 0
