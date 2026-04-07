@echo off
setlocal EnableExtensions

set "ROOT=C:\dev\DiagnosticoStressWindows"
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%build-run-viewer.ps1"
set "OUT_HTML=%SCRIPT_DIR%viewer-report.html"

if not exist "%PS_SCRIPT%" (
  echo [FALHA] build-run-viewer.ps1 nao encontrado.
  pause
  exit /b 1
)

echo.
echo [PROCESSANDO] Gerando viewer a partir de:
echo %ROOT%
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -RootPath "%ROOT%" -OutputPath "%OUT_HTML%"
if errorlevel 1 (
  echo.
  echo [FALHA] Nao foi possivel gerar o viewer.
  pause
  exit /b 1
)

echo.
echo [OK] Viewer gerado em:
echo %OUT_HTML%
echo.

start "" "%OUT_HTML%"
