@echo off
title Extrator de Dados RDO - REPAR
color 0A

echo.
echo =====================================================
echo        EXTRATOR DE DADOS RDO - REPAR
echo =====================================================
echo.

REM Verifica se o Python esta instalado
echo Verificando instalacao do Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERRO: Python nao foi encontrado no sistema!
    echo.
    echo Por favor, instale o Python primeiro:
    echo    - Acesse: https://python.org/downloads
    echo    - Baixe a versao mais recente
    echo    - Durante a instalacao, marque "Add to PATH"
    echo.
    pause
    exit /b 1
)

echo Python encontrado!
python --version

REM Verifica se o arquivo Python existe
echo.
echo Verificando arquivo do script...
if not exist "extrator_rdo.py" (
    echo ERRO: Arquivo 'extrator_rdo.py' nao encontrado!
    echo Certifique-se de que ambos os arquivos estao na mesma pasta.
    echo.
    pause
    exit /b 1
)
echo Script encontrado!

REM Instala/atualiza as dependencias necessarias
echo.
echo Instalando/verificando dependencias...
echo    - pandas
echo    - openpyxl
pip install pandas openpyxl --quiet --upgrade
if errorlevel 1 (
    echo Aviso: Possivel problema na instalacao das dependencias
    echo O script tentara continuar...
)

REM Executa o script Python
echo.
echo Iniciando extrator de dados...
echo.
python extrator_rdo.py

REM Verifica o resultado da execucao
if errorlevel 1 (
    echo.
    echo ERRO durante a execucao do script!
    echo Verifique o arquivo 'extrator_rdo.log' para detalhes
    echo.
) else (
    echo.
    echo Script executado!
    echo Verifique a pasta de destino para o arquivo resultado
)

echo.
echo =====================================================
pause