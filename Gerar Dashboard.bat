@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Gerar Dashboard SCM

echo ============================================================
echo    Gerador de Dashboard SCM Semanal
echo ------------------------------------------------------------
echo    Uma janela vai abrir para voce escolher a planilha.
echo    (ou arraste o arquivo .xlsx para cima deste .bat)
echo ============================================================
echo.

set "PY=python"
where python >nul 2>nul || set "PY=py"

%PY% "gerar_dashboard.py" %*

if errorlevel 1 (
  echo.
  echo ------------------------------------------------------------
  echo  Ocorreu um erro, ou o Python nao foi encontrado.
  echo  Verifique se o Python esta instalado e tente novamente.
  echo ------------------------------------------------------------
  pause
)
