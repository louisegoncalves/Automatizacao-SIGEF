@echo off
title Robô de Prestação de Contas v3.1 - by Louise <3

cls
color 0F

echo.
echo      db    db    db    db    db    db    db    db    db    db
echo    .d88b .d88b .d88b .d88b .d88b .d88b .d88b .d88b .d88b .d88b
echo    8888888888888888888888888888888888888888888888888888888888
echo    `Y8888888888888888888888888888888888888888888888888888888P'
echo      `Y888888888888888888888888888888888888888888888888888P'
echo        `Y88888888888888888888888888888888888888888888888P'
echo          `Y8888888888888888888888888888888888888888888P'
echo            `Y888888888888888888888888888888888888888P'
echo              `Y88888888888888888888888888888888888P'
echo                `Y8888888888888888888888888888888P'
echo                  `Y888888888888888888888888888P'
echo                    `Y88888888888888888888888P'
echo                      `Y8888888888888888888P'
echo                        `Y888888888888888P'
echo                          `Y88888888888P'
echo                            `Y8888888P'
echo                              `Y888P'
echo                                `Y'
echo.
echo           BEM-VINDO AO ROBO DE PRESTACAO DE CONTAS!
echo                     Feito com ^<3 por Louise
echo.

timeout /t 5 /nobreak >nul

cls
color 0F

echo.
echo      +-------------------------------------------------+
echo      ^|  VERIFICANDO AMBIENTE DO ROBO...                ^|
echo      +-------------------------------------------------+
echo.

wmic process where "name='chrome.exe' and commandline like '%%--remote-debugging-port=9222%%'" get ProcessID | findstr /R "[0-9]" >nul

if %errorlevel%==0 (
    echo  [ SUCESSO ] Janela de depuracao do Chrome ja esta aberta.
    echo  Pulando a etapa de inicializacao do navegador.
    echo.
    timeout /t 3 /nobreak >nul
    goto IniciarRobo
) else (
    echo  [ AVISO ] Janela de depuracao nao encontrada.
    echo  Iniciando uma nova sessao do Chrome...
    echo.
    timeout /t 3 /nobreak >nul
    goto AbrirChrome
)


:AbrirChrome

cls
color 0F

echo.
echo           +---------------------------------------------+
echo           ^|      PASSO 1: INICIAR O NAVEGADOR            ^|
echo           +---------------------------------------------+
echo.
echo    O Google Chrome sera iniciado em modo de depuracao. Feche todas as instancias do Chrome antes de prosseguir.
echo.

pause
start "Chrome Debug" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebugProfile"

timeout /t 3 /nobreak >nul

echo           +---------------------------------------------+
echo           ^|      PASSO 2: LOGIN MANUAL NO SIGEF          ^|
echo           +---------------------------------------------+
echo.
echo    IMPORTANTE: Faca o login no SIGEF e resolva o CAPTCHA.
echo    Apos o login, volte para esta janela.
echo.
echo.

pause
goto IniciarRobo

:IniciarRobo

cls
color 0F

echo           +---------------------------------------------+
echo           ^|      EXECUTANDO O ROBO DE AUTOMACAO...       ^|
echo           +---------------------------------------------+
echo.
echo    Tudo pronto! Iniciando o script Python.
echo    Preencha as informacoes corretamente: unidade gestora, empenho, exercicio financeiro...
echo.
echo    Voce pode acompanhar o progresso na janela do Chrome.
echo.

cd /d "C:\Users\SESDEC-CAF\Desktop\Automacao\Robozinho"

python "Realizar Prestacao de Contas.py"

cls
color 0F

echo.
echo.
echo              +-------------------------------------+
echo              ^|        OPERACAO FINALIZADA!         ^|
echo              +-------------------------------------+
echo.
echo.
echo            Por: LOUISE-SESDEC
echo.
echo.

pause