@echo off
setlocal enabledelayedexpansion

echo ====================================================
echo   Gerador de Executavel - CobrancaNF
echo ====================================================
echo.

:: Verifica se o Python esta instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado. Por favor, instale o Python antes de continuar.
    pause
    exit /b
)

:: Verifica/Instala dependencias do requirements.txt
echo [INFO] Verificando dependencias...
pip install -r requirements.txt

:: Limpa pastas de build antigas
if exist build (
    echo [INFO] Removendo pasta build antiga...
    rd /s /q build
)
if exist dist (
    echo [INFO] Removendo pasta dist antiga...
    rd /s /q dist
)

:: Gera o executavel
echo.
echo [INFO] Gerando executavel a partir de main.pyw...
echo.

:: O PyInstaller ira coletar as dependencias do src automaticamente
:: Adicionamos --noconsole para interface GUI
pyinstaller --noconfirm --onefile --windowed ^
    --name "CobrancaNF" ^
    --clean ^
    "main.pyw"

echo.
if %errorlevel% equ 0 (
    echo ====================================================
    echo   SUCESSO! O executavel foi gerado na pasta 'dist'
    echo   Nome: CobrancaNF.exe
    echo ====================================================
) else (
    echo [ERRO] Ocorreu um problema ao gerar o executavel.
)

pause
