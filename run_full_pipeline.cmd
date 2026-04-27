@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0"
pushd "%REPO_ROOT%"
set "EXIT_CODE=0"

set "ORIGINAL_CODE_FOR_AUTOTEST_PAUSE_ON_ERROR=%CODE_FOR_AUTOTEST_PAUSE_ON_ERROR%"
set "CODE_FOR_AUTOTEST_PAUSE_ON_ERROR=0"
call "%REPO_ROOT%initialize_environment.cmd"
if defined ORIGINAL_CODE_FOR_AUTOTEST_PAUSE_ON_ERROR (
    set "CODE_FOR_AUTOTEST_PAUSE_ON_ERROR=%ORIGINAL_CODE_FOR_AUTOTEST_PAUSE_ON_ERROR%"
) else (
    set "CODE_FOR_AUTOTEST_PAUSE_ON_ERROR="
)
if errorlevel 1 (
    echo Environment initialization failed. The full pipeline will not start.
    set "EXIT_CODE=1"
    goto :finish
)

set "PYTHON_LAUNCHER="
python -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=python"
if not defined PYTHON_LAUNCHER py -3 -c "import sys" >nul 2>nul && set "PYTHON_LAUNCHER=py -3"

if not defined PYTHON_LAUNCHER (
    echo Python was not found after initialization.
    set "EXIT_CODE=1"
    goto :finish
)

call %PYTHON_LAUNCHER% "%REPO_ROOT%full_test_pipeline.py" %*
set "EXIT_CODE=%ERRORLEVEL%"

:finish
if not "%EXIT_CODE%"=="0" (
    echo.
    echo run_full_pipeline.cmd failed with exit code %EXIT_CODE%.
    if not "%CODE_FOR_AUTOTEST_PAUSE_ON_ERROR%"=="0" (
        echo The window will stay open so you can review the traceback above.
        pause
    )
)
popd
exit /b %EXIT_CODE%
