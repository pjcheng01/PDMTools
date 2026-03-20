@echo off
setlocal

REM Move to script directory (project root)
cd /d "%~dp0"

set "PROJECT_FILE=%~dp0PDMTools\PDMTools.csproj"
set "OUTPUT_DIR=%~dp0PDMTools\bin\Release\net8.0-windows"

echo [INFO] Building Release...

REM Prefer user-local dotnet install, fallback to PATH
if exist "%USERPROFILE%\.dotnet\dotnet.exe" (
  "%USERPROFILE%\.dotnet\dotnet.exe" build "%PROJECT_FILE%" -c Release
) else (
  dotnet build "%PROJECT_FILE%" -c Release
)

if errorlevel 1 (
  echo.
  echo [ERROR] Build failed. Please check messages above.
  pause
  exit /b 1
)

echo.
echo [INFO] Build succeeded.

if exist "%OUTPUT_DIR%\PDMTools.exe" (
  echo [INFO] Opening output folder: "%OUTPUT_DIR%"
  start "" "%OUTPUT_DIR%"
) else (
  echo [WARN] Build finished but PDMTools.exe was not found at:
  echo        "%OUTPUT_DIR%"
  pause
  exit /b 2
)

echo [DONE]
endlocal
