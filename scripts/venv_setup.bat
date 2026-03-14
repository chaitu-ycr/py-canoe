@echo off

REM ============================
REM Create or Update Python Virtual Environment and Install Dependencies using uv
REM ============================

title Creating/Updating Tool Environment with uv...

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

popd
goto :EOF

REM ----------------------------
REM Error Handler
REM ----------------------------
:ERROR
echo Failed to set up virtual environment and install dependencies due to error %ERRORLEVEL%.
popd
pause
goto :EOF