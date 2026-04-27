@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0"
pushd "%REPO_ROOT%"
set "EXIT_CODE=0"

set "PYTHON_LAUNCHER="
python -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=python"
if not defined PYTHON_LAUNCHER py -3 -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=py -3"

if not defined PYTHON_LAUNCHER (
    echo Python was not found. Installing Python via winget...
    where winget >nul 2>nul || (
        echo winget is not available. Install Python manually and rerun this script.
        set "EXIT_CODE=1"
        goto :finish
    )
    winget install -e --id Python.Python.3.13 --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo Python installation failed.
        set "EXIT_CODE=1"
        goto :finish
    )
    python -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=python"
    if not defined PYTHON_LAUNCHER py -3 -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=py -3"
)

if not defined PYTHON_LAUNCHER (
    echo Python was installed but is not visible in the current shell. Open a new terminal and rerun this script.
    set "EXIT_CODE=1"
    goto :finish
)

set "REQUIREMENTS_FILE=%REPO_ROOT%requirements.txt"
call %PYTHON_LAUNCHER% -m pip --version >nul 2>nul
if errorlevel 1 (
    echo pip is not available. Bootstrapping it with ensurepip...
    call %PYTHON_LAUNCHER% -m ensurepip --upgrade
    if errorlevel 1 (
        echo Failed to bootstrap pip.
        set "EXIT_CODE=1"
        goto :finish
    )
)

call %PYTHON_LAUNCHER% -c "import importlib.util, sys; required=('psutil','openpyxl','pywinauto','pyautogui','pytest'); missing=[name for name in required if importlib.util.find_spec(name) is None]; print('Missing Python packages: ' + ', '.join(missing) if missing else 'All required Python packages are already installed.'); raise SystemExit(1 if missing else 0)"
if errorlevel 1 (
    echo Installing Python requirements from "%REQUIREMENTS_FILE%"...
    call %PYTHON_LAUNCHER% -m pip install -r "%REQUIREMENTS_FILE%"
    if errorlevel 1 (
        echo Python package installation failed.
        set "EXIT_CODE=1"
        goto :finish
    )
)

set "INIT_ARTIFACTS_ROOT=%REPO_ROOT%results\initialize_environment_runs"
echo Initialization artifacts will be written under "%INIT_ARTIFACTS_ROOT%".
call %PYTHON_LAUNCHER% "%REPO_ROOT%initialize_environment.py" --artifacts-root "%INIT_ARTIFACTS_ROOT%" %*
set "EXIT_CODE=%ERRORLEVEL%"

:finish
if not "%EXIT_CODE%"=="0" (
    echo.
    echo initialize_environment.cmd failed with exit code %EXIT_CODE%.
    if not "%CODE_FOR_AUTOTEST_PAUSE_ON_ERROR%"=="0" (
        echo The window will stay open so you can review the traceback above.
        pause
    )
)
popd
exit /b %EXIT_CODE%
