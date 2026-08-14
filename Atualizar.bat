@echo off
setlocal enabledelayedexpansion

:: ====================================================================
::  ATUALIZAR / ABRIR ANALISADOR DE NF-e HEC (via servidor, com copia
::  local automatica)
:: ====================================================================
::  Copia a pasta do programa do servidor para o computador do usuario
::  (so o que mudou - da segunda vez em diante e quase instantaneo) e
::  abre a COPIA LOCAL, ja atualizada. Assim o usuario nunca precisa
::  entrar na pasta do servidor nem abrir o executavel direto da rede
::  (rodar direto da rede e sempre lento e mais instavel).
::
::  CONFIGURACAO (uma vez so, por computador):
::  O caminho da pasta do servidor NAO fica gravado neste arquivo (ele
::  e enviado ao GitHub e o repositorio e publico) -- fica salvo num
::  arquivo local "servidor.txt", do lado deste .bat, criado
::  automaticamente na primeira vez que voce roda o programa (o
::  script pergunta o caminho e grava sozinho para as proximas vezes).
:: ====================================================================

cd /d "%~dp0"
set CONFIG=%~dp0servidor.txt
set DESTINO=%LOCALAPPDATA%\AnalisadorNFeHEC

if exist "%CONFIG%" (
    set /p ORIGEM=<"%CONFIG%"
) else (
    echo ====================================================================
    echo   Primeira vez rodando neste computador -- configuracao do servidor
    echo ====================================================================
    echo.
    echo Informe o caminho da pasta do SERVIDOR onde fica o programa
    echo ^(exemplo: \\NOME_DO_SERVIDOR\pasta\subpasta\ANALISADOR_XML^)
    echo.
    set /p ORIGEM="Caminho do servidor: "
    if "!ORIGEM!"=="" (
        echo.
        echo       Nenhum caminho informado. Cancelado.
        pause
        exit /b 1
    )
    >"%CONFIG%" echo !ORIGEM!
    echo.
    echo Caminho salvo em "servidor.txt" -- nas proximas vezes nao pergunta de novo.
    echo.
)

echo ====================================================================
echo   ANALISADOR DE NF-e HEC - Atualizando...
echo ====================================================================
echo.
echo Pasta do servidor - ORIGEM: !ORIGEM!
echo Pasta local - DESTINO:      %DESTINO%
echo.

if not exist "!ORIGEM!" (
    echo --------------------------------------------------------------
    echo   ERRO: nao encontrei a pasta do servidor acima.
    echo.
    echo   Possiveis causas:
    echo   1. O caminho salvo em "servidor.txt" esta errado. Apague o
    echo      arquivo "servidor.txt" desta pasta e rode este .bat de
    echo      novo para informar o caminho correto.
    echo   2. Este computador nao tem acesso a essa pasta da rede.
    echo --------------------------------------------------------------
    echo.
    pause
    exit /b 1
)

echo Sincronizando com o servidor - so baixa o que mudou...
echo Na primeira vez pode demorar 1-2 minutos. Nas proximas, segundos.
echo.
robocopy "!ORIGEM!" "%DESTINO%" /MIR /XF Atualizar.bat servidor.txt /NFL /NDL /NJH /NJS /NP /R:1 /W:1
set RC=%errorlevel%
echo.
echo Codigo de retorno do robocopy: %RC%  -- de 0 a 7 e normal/sucesso

if %RC% GEQ 8 (
    echo --------------------------------------------------------------
    echo   ERRO: a copia do servidor falhou. Confira se voce tem
    echo   permissao de leitura na pasta do servidor e tente de novo.
    echo --------------------------------------------------------------
    echo.
    pause
    exit /b 1
)

set ACHOU=
for /r "%DESTINO%" %%F in (Analisador_NFe_HEC_*.exe) do set ACHOU=%%F

if "%ACHOU%"=="" (
    echo --------------------------------------------------------------
    echo   ERRO: a copia funcionou, mas nao encontrei nenhum arquivo
    echo   Analisador_NFe_HEC_*.exe dentro de:
    echo   %DESTINO%
    echo.
    echo   Confira se a pasta do servidor - ORIGEM - realmente contem a
    echo   pasta do programa, exemplo: Analisador_NFe_HEC_v01.0 com o
    echo   .exe e a pasta _internal dentro.
    echo --------------------------------------------------------------
    echo.
    pause
    exit /b 1
)

echo.
echo Arquivo encontrado: %ACHOU%
echo Tentando abrir...
start "" "%ACHOU%"

echo Aguardando 5 segundos para conferir se o programa realmente abriu...
timeout /t 5 >nul

tasklist /fi "imagename eq Analisador_NFe_HEC*" 2>nul | find /i "Analisador_NFe_HEC" >nul
if errorlevel 1 (
    echo --------------------------------------------------------------
    echo   ATENCAO: o comando de abrir foi enviado, mas o programa NAO
    echo   aparece rodando agora. Causas mais comuns:
    echo.
    echo   1. O Windows SmartScreen bloqueou silenciosamente o arquivo
    echo      por ele ter acabado de ser copiado da rede - comum na
    echo      primeira vez. Va ate a pasta abaixo e abra o .exe com
    echo      duplo-clique manualmente uma vez -- se aparecer um aviso
    echo      azul "O Windows protegeu o computador", clique em
    echo      "Mais informacoes" e depois em "Executar assim mesmo".
    echo      Depois disso, o atalho automatico deve passar a funcionar.
    echo.
    echo   2. O antivirus da empresa pode ter colocado o arquivo em
    echo      quarentena. Confira o historico/quarentena do antivirus.
    echo.
    echo   Pasta onde o programa foi copiado:
    echo   %DESTINO%
    echo --------------------------------------------------------------
) else (
    echo --------------------------------------------------------------
    echo   OK! O Analisador de NF-e HEC esta rodando.
    echo --------------------------------------------------------------
)
echo.
pause
