@echo off
setlocal
set "REPO_ROOT=%~dp0"
pushd "%REPO_ROOT%"

call "%REPO_ROOT%initialize_environment.cmd"
if errorlevel 1 (
    echo Environment initialization failed. The full pipeline will not start.
    popd
    exit /b 1
)

set "PYTHON_LAUNCHER="
python -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=python"
if not defined PYTHON_LAUNCHER py -3 -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=py -3"

if not defined PYTHON_LAUNCHER (
    echo Python was not found after initialization.
    popd
    exit /b 1
)

call %PYTHON_LAUNCHER% "%REPO_ROOT%full_test_pipeline.py" %*
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
