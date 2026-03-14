@echo off

title "uploading package to pypi"

REM Save original directory and move to project root
set "ORIGIN_DIR=%CD%"
pushd %~dp0
cd ..

REM ----------------------------
REM 1. Check if uv is installed, else install it
REM ----------------------------
:CHECK_UV
where uv >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo uv is already installed.
    uv --version
) else (
    echo uv not found. Installing uv...
    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    if %ERRORLEVEL% NEQ 0 goto ERROR
    echo uv installation completed.
)

REM ----------------------------
REM 2. Sync Dependencies
REM ----------------------------
echo Syncing dependencies with uv...
uv sync --link-mode=copy
if %ERRORLEVEL% NEQ 0 goto ERROR
echo Completed syncing dependencies.

REM ----------------------------
REM 3. Publish Package to PyPI
REM ----------------------------
echo Publishing package to PyPI...
uv publish
if %ERRORLEVEL% NEQ 0 goto ERROR
echo Package published successfully.

:END
cd %ORIGIN_DIR%
pause
GOTO :eof

:ERROR
title "Failed to upload package to PyPI due to error %ERRORLEVEL%"
cd %ORIGIN_DIR%
pause
GOTO :eof
