@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>&1
if errorlevel 1 (
  echo [ERRO] Python nao encontrado. Instale Python 3.11 ou superior e marque "Add Python to PATH".
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Criando ambiente virtual...
  py -m venv .venv
  if errorlevel 1 goto :erro
)

echo Instalando dependencias...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :erro

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :erro

if not exist "config.yaml" (
  copy /Y "config.example.yaml" "config.yaml" >nul
  echo Criado config.yaml a partir do exemplo.
)

if not exist "data\lojas.csv" (
  if not exist "data" mkdir "data"
  copy /Y "data\lojas.exemplo.csv" "data\lojas.csv" >nul
  echo Criado data\lojas.csv a partir do exemplo.
)

echo.
echo Instalacao concluida.
echo Edite config.yaml e data\lojas.csv antes da primeira execucao.
exit /b 0

:erro
echo.
echo [ERRO] Falha durante a instalacao.
pause
exit /b 1
