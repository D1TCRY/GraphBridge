@echo off
setlocal enableextensions

cd /d "%~dp0"

rem Safe default: build and verify only. Uploads require both an explicit mode
rem and an exact environment confirmation. Twine obtains credentials from its
rem normal secure configuration; this script never accepts or stores them.
set "MODE=%~1"
if "%~1"=="" set "MODE=build"

if /i "%MODE%"=="build" goto :validated
if /i "%MODE%"=="test" goto :validated
if /i "%MODE%"=="pypi" goto :validated
echo [ERROR] Invalid argument: %MODE%
echo Usage: publish.bat [build^|test^|pypi]
exit /b 2

:validated
where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not available in PATH.
    exit /b 1
)

python -c "import build, twine" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Install release tools explicitly: python -m pip install build twine
    exit /b 1
)

echo [1/5] Unit tests...
python -m pytest tests\unit --cov=graphbridge --cov-branch --cov-report=term-missing
if errorlevel 1 exit /b 1

echo [2/5] Ruff...
python -m ruff check .
if errorlevel 1 exit /b 1

echo [3/5] mypy...
python -m mypy
if errorlevel 1 exit /b 1

echo [4/5] Clean build and package verification...
if exist "%CD%\build" rmdir /s /q "%CD%\build"
if exist "%CD%\dist" rmdir /s /q "%CD%\dist"
python -m build
if errorlevel 1 exit /b 1
python -m twine check "dist\*"
if errorlevel 1 exit /b 1
if exist "%CD%\build" rmdir /s /q "%CD%\build"

if /i "%MODE%"=="build" (
    echo [5/5] Build verified. No package was uploaded.
    exit /b 0
)

if /i "%MODE%"=="test" (
    if not "%GRAPHBRIDGE_PUBLISH_CONFIRM%"=="TESTPYPI" (
        echo [ERROR] Set GRAPHBRIDGE_PUBLISH_CONFIRM=TESTPYPI to authorize this upload.
        exit /b 3
    )
    echo [5/5] Uploading to TestPyPI...
    python -m twine upload --repository testpypi "dist\*"
    exit /b %errorlevel%
)

if not "%GRAPHBRIDGE_PUBLISH_CONFIRM%"=="PYPI" (
    echo [ERROR] Set GRAPHBRIDGE_PUBLISH_CONFIRM=PYPI to authorize this upload.
    exit /b 3
)
echo [5/5] Uploading to PyPI...
python -m twine upload "dist\*"
exit /b %errorlevel%
