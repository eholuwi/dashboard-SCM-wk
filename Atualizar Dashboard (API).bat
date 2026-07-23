@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Atualizar Dashboard SCM (API)

echo ============================================================
echo    Dashboard SCM - coleta automatica pela API
echo ------------------------------------------------------------
echo    Nao e preciso exportar nenhuma planilha.
echo    Os dados vem direto do SCM (mansrvapp03:5715).
echo ============================================================
echo.

set "PY=python"
where python >nul 2>nul || set "PY=py"

%PY% "coletar.py" %*

if errorlevel 1 (
  echo.
  echo ------------------------------------------------------------
  echo  Falhou. Possiveis causas:
  echo    - sem rede/VPN ate mansrvapp03:5715
  echo    - Python nao instalado
  echo  Alternativa: use "Gerar Dashboard.bat" com as planilhas
  echo  exportadas a mao (fluxo antigo, continua funcionando).
  echo ------------------------------------------------------------
  pause
  exit /b 1
)

REM ---- Publicacao na pasta de rede (opcional) ----
REM Preencha PUBLICAR com o caminho da pasta compartilhada e descomente
REM as linhas abaixo para que cada coleta ja publique para os compradores.
REM set "PUBLICAR=\\servidor\Compras\DashboardSCM"
if defined PUBLICAR (
  echo Publicando em %PUBLICAR% ...
  if not exist "%PUBLICAR%" mkdir "%PUBLICAR%"
  for /f "delims=" %%F in ('dir /b /o-d "WK\Dashboard SCM WK*.html"') do (
    copy /y "WK\%%F" "%PUBLICAR%\Dashboard SCM.html" >nul
    goto :publicado
  )
  :publicado
  echo Publicado.
)
