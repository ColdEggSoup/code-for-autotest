@echo off
setlocal
set "REPO_ROOT=%~dp0"
pushd "%REPO_ROOT%"

set "PYTHON_LAUNCHER="
python -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=python"
if not defined PYTHON_LAUNCHER py -3 -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=py -3"

if not defined PYTHON_LAUNCHER (
    echo Python was not found. Installing Python via winget...
    where winget >nul 2>nul || (
        echo winget is not available. Install Python manually and rerun this script.
        popd
        exit /b 1
    )
    winget install -e --id Python.Python.3.13 --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo Python installation failed.
        popd
        exit /b 1
    )
    python -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=python"
    if not defined PYTHON_LAUNCHER py -3 -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=py -3"
)

if not defined PYTHON_LAUNCHER (
    echo Python was installed but is not visible in the current shell. Open a new terminal and rerun this script.
    popd
    exit /b 1
)

call %PYTHON_LAUNCHER% "%REPO_ROOT%initialize_environment.py" %*
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
