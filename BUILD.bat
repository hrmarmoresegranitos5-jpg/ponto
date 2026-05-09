@echo off
chcp 65001 >nul
cls

echo.
echo ================================================
echo   Sistema de Ponto v5.1 — Build para Windows
echo ================================================
echo.

:: ── Verificar Python ────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado.
    echo.
    echo  Instale em: https://python.org/downloads
    echo  IMPORTANTE: marque a opcao "Add Python to PATH"
    echo.
    pause
    exit /b 1
)
echo [OK] Python encontrado

:: ── Verificar NSIS ──────────────────────────────────────────
makensis /VERSION >nul 2>&1
if errorlevel 1 (
    echo [AVISO] makensis nao encontrado no PATH.
    echo.
    echo  Adicione o NSIS ao PATH ou instale em:
    echo  https://nsis.sourceforge.io/Download
    echo.
    echo  Alternativamente, compile installer.nsi manualmente
    echo  clicando com botao direito > "Compile NSIS Script"
    echo.
)

:: ── Instalar dependências ────────────────────────────────────
echo.
echo Instalando dependencias Python...
pip install --quiet pyinstaller openpyxl xlrd reportlab requests
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas

:: ── Gerar ícone ─────────────────────────────────────────────
echo.
if not exist icone.ico (
    echo Gerando icone...
    python gerar_icone.py
    if errorlevel 1 (
        echo [AVISO] Falha ao gerar icone. Continuando sem icone...
    )
)

:: ── Limpar builds anteriores ────────────────────────────────
echo.
echo Limpando builds anteriores...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

:: ── PyInstaller ─────────────────────────────────────────────
echo.
echo Compilando com PyInstaller...
echo (isso pode levar 1-3 minutos)
echo.
pyinstaller ponto.spec --noconfirm
if errorlevel 1 (
    echo.
    echo [ERRO] PyInstaller falhou. Veja o log acima.
    pause
    exit /b 1
)
echo.
echo [OK] dist\SistemaPonto\ criado com sucesso

:: ── NSIS ────────────────────────────────────────────────────
echo.
echo Criando instalador com NSIS...
makensis installer.nsi
if errorlevel 1 (
    echo.
    echo [AVISO] makensis falhou ou nao foi encontrado.
    echo         Compile installer.nsi manualmente no NSIS.
    goto :sem_instalador
)
echo [OK] SistemaPonto_Setup.exe criado!

:: ── Sucesso ─────────────────────────────────────────────────
echo.
echo ================================================
echo   BUILD CONCLUIDO COM SUCESSO!
echo ================================================
echo.
echo   Arquivo: SistemaPonto_Setup.exe
echo.
echo   O instalador:
echo    OK  Instala com duplo clique
echo    OK  Cria atalho no Desktop e Menu Iniciar
echo    OK  Abre sem terminal, sem Python visivel
echo    OK  Aparece no Painel de Controle
echo.
goto :fim

:sem_instalador
echo.
echo ================================================
echo   PyInstaller OK — Instalador precisa ser
echo   compilado manualmente (veja instrucoes acima)
echo ================================================
echo.

:fim
pause
