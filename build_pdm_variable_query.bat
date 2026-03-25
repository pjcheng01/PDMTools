@echo off
setlocal
cd /d "%~dp0"

if exist "%USERPROFILE%\.dotnet\dotnet.exe" (
  "%USERPROFILE%\.dotnet\dotnet.exe" build "%~dp0PdmVariableQuery\PdmVariableQuery.csproj" -c Release
) else (
  dotnet build "%~dp0PdmVariableQuery\PdmVariableQuery.csproj" -c Release
)

if errorlevel 1 (
  echo [ERROR] Build failed.
  pause
  exit /b 1
)

set "OUT=%~dp0PdmVariableQuery\bin\Release\net48"
if exist "%OUT%\PdmVariableQuery.exe" (
  echo [OK] Output: "%OUT%\PdmVariableQuery.exe"
  start "" "%OUT%"
) else (
  echo [WARN] Expected exe not found under: "%OUT%"
)

endlocal
