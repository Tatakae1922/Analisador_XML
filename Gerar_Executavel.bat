@echo off
setlocal enabledelayedexpansion

:: ====================================================================
::  GERAR EXECUTAVEL -- Analisador de NF-e HEC
:: ====================================================================
::  Script de uso do DESENVOLVEDOR (nao e o Atualizar.bat que o usuario
::  final roda). Automatiza a geracao de um novo executavel depois de
::  alterar analisador_xml.py:
::
::    1. Le a versao automaticamente de dentro do proprio .py (nao
::       precisa digitar o numero da versao na mao)
::    2. Limpa a geracao anterior (build/dist/*.spec)
::    3. Roda o PyInstaller com os parametros corretos (onedir,
::       ícone e logo embutidos)
::    4. Abre o executavel gerado, para voce conferir rapidamente
::    5. Abre a pasta pronta no Explorer -- e so arrastar essa pasta
::       para a pasta do servidor (a mesma que o Atualizar.bat dos
::       usuarios sincroniza)
:: ====================================================================

cd /d "%~dp0"

set VERSAO=
for /f "usebackq delims=" %%L in (`findstr /b "VERSAO_PROGRAMA" analisador_xml.py`) do set LINHA=%%L
for /f "tokens=2 delims==" %%V in ("!LINHA!") do set VERSAO_BRUTA=%%V
for /f "tokens=1 delims=#" %%V in ("!VERSAO_BRUTA!") do set VERSAO_BRUTA=%%V
set VERSAO=!VERSAO_BRUTA: =!
set VERSAO=!VERSAO:"=!

if "!VERSAO!"=="" (
    echo --------------------------------------------------------------
    echo   ERRO: nao consegui ler a linha "VERSAO_PROGRAMA = ..." dentro
    echo   de analisador_xml.py. Confira se o arquivo esta na mesma
    echo   pasta deste .bat e se a linha nao foi renomeada.
    echo --------------------------------------------------------------
    pause
    exit /b 1
)

set NOME=Analisador_NFe_HEC_!VERSAO!

echo ====================================================================
echo   Gerando !NOME! ...
echo ====================================================================
echo.

echo [1/4] Limpando geracao anterior (build, dist, *.spec)...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "!NOME!.spec" del "!NOME!.spec"
echo       OK.
echo.

echo [2/4] Rodando o PyInstaller (pode levar 1-2 minutos)...
python -m PyInstaller --onedir --windowed --name "!NOME!" --icon logo_hec.ico --add-data "logo.png;." --add-data "logo_hec.ico;." analisador_xml.py

if not exist "dist\!NOME!\!NOME!.exe" (
    echo.
    echo ====================================================================
    echo   ERRO: o executavel nao foi gerado. Confira as mensagens do
    echo   PyInstaller acima.
    echo ====================================================================
    pause
    exit /b 1
)
echo       OK. Gerado em dist\!NOME!\!NOME!.exe
echo.

echo [3/4] Abrindo o programa para voce conferir antes de distribuir...
start "" "dist\!NOME!\!NOME!.exe"
echo       Teste o programa (abra uma pasta de XML de teste, gere uma
echo       planilha). Feche a janela dele quando terminar de conferir.
echo.

echo [4/4] Abrindo a pasta pronta no Explorer...
echo       Copie a pasta "!NOME!" (a pasta INTEIRA, com o .exe e a
echo       _internal dentro) para a pasta do servidor que o Atualizar.bat
echo       dos usuarios sincroniza -- substituindo a pasta da versao
echo       anterior, se houver.
start "" explorer.exe "%~dp0dist"

echo.
echo ====================================================================
echo   CONCLUIDO. Pasta pronta em: dist\!NOME!
echo   Falta so: copiar essa pasta para o servidor.
echo   (Lembrete: se o codigo mudou, faca "git add", "git commit" e
echo   "git push" tambem -- este script so gera o executavel.)
echo ====================================================================
pause
