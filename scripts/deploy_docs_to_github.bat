@echo off

REM ============================
REM Deploy MkDocs Documentation to GitHub Pages
REM ============================

REM Set window title
title Deploying Documentation to GitHub Pages

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
REM 3. Deploy Documentation to GitHub Pages
REM ----------------------------
echo Deploying documentation to GitHub Pages...
uv run mkdocs gh-deploy
if %ERRORLEVEL% NEQ 0 goto ERROR
echo Documentation deployed successfully.
popd
cd "%ORIGIN_DIR%"
goto :EOF

REM ----------------------------
REM Error Handler
REM ----------------------------
:ERROR
echo Failed to deploy documentation due to error %ERRORLEVEL%
popd
cd "%ORIGIN_DIR%"
pause
goto :EOF