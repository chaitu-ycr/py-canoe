@echo off

REM ============================
REM Build Python Wheel Package
REM ============================

REM Set window title
title Building Python Wheel Package

REM Move to script directory and then project root
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
REM 3. Build the Wheel Package
REM ----------------------------
uv build
if %ERRORLEVEL% NEQ 0 goto ERROR
echo wheel package built successfully.

popd
goto :EOF

REM ----------------------------
REM Error Handler
REM ----------------------------
:ERROR
title Failed to build wheel package due to error %ERRORLEVEL%
popd
pause
goto :EOF
