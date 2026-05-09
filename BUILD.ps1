# ============================================================
#  BUILD.ps1 — Script de Build Completo  •  Sistema de Ponto
#  Execute com PowerShell no Windows:
#    .\BUILD.ps1
# ============================================================

$ErrorActionPreference = "Stop"
$APP_NAME    = "Sistema de Ponto"
$APP_VERSION = "5.1"

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  $APP_NAME v$APP_VERSION — Build para Windows" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# ── Pré-requisitos ──────────────────────────────────────────

function Check-Command($cmd, $install_hint) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        Write-Host "  [ERRO] '$cmd' não encontrado." -ForegroundColor Red
        Write-Host "  → $install_hint" -ForegroundColor Yellow
        exit 1
    }
    Write-Host "  [OK] $cmd" -ForegroundColor Green
}

Write-Host "Verificando pré-requisitos..." -ForegroundColor White
Check-Command "python"     "Instale em https://python.org (marque 'Add to PATH')"
Check-Command "pip"        "Reinstale o Python marcando 'Add to PATH'"
Check-Command "makensis"   "Instale o NSIS em https://nsis.sourceforge.io"

Write-Host ""

# ── Instalar dependências Python ────────────────────────────

Write-Host "Instalando dependências Python..." -ForegroundColor White
pip install --quiet pyinstaller openpyxl xlrd reportlab requests
if ($LASTEXITCODE -ne 0) { Write-Host "Falha no pip install" -ForegroundColor Red; exit 1 }
Write-Host "  [OK] Dependências instaladas" -ForegroundColor Green
Write-Host ""

# ── Gerar ícone se não existir ──────────────────────────────

if (-not (Test-Path "icone.ico")) {
    Write-Host "Gerando ícone padrão..." -ForegroundColor White
    python gerar_icone.py
    Write-Host "  [OK] icone.ico criado" -ForegroundColor Green
    Write-Host ""
}

# ── PyInstaller ─────────────────────────────────────────────

Write-Host "Compilando com PyInstaller..." -ForegroundColor White
Write-Host "  (isso pode levar 1-3 minutos)" -ForegroundColor Gray

# Limpa builds anteriores
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist")  { Remove-Item -Recurse -Force "dist"  }

pyinstaller ponto.spec --noconfirm
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "  [ERRO] PyInstaller falhou. Veja o log acima." -ForegroundColor Red
    exit 1
}

Write-Host "  [OK] dist\SistemaPonto\ criado" -ForegroundColor Green
Write-Host ""

# ── Verificar tamanho do build ──────────────────────────────

$size = (Get-ChildItem "dist\SistemaPonto" -Recurse | Measure-Object -Property Length -Sum).Sum
$sizeMB = [math]::Round($size / 1MB, 1)
Write-Host "  Tamanho do pacote: $sizeMB MB" -ForegroundColor Gray
Write-Host ""

# ── NSIS — Criar instalador ─────────────────────────────────

Write-Host "Criando instalador com NSIS..." -ForegroundColor White
makensis installer.nsi
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "  [ERRO] NSIS falhou. Veja o log acima." -ForegroundColor Red
    exit 1
}

Write-Host "  [OK] SistemaPonto_Setup.exe criado" -ForegroundColor Green
Write-Host ""

# ── Resultado ───────────────────────────────────────────────

$setupSize = [math]::Round((Get-Item "SistemaPonto_Setup.exe").Length / 1MB, 1)

Write-Host "================================================" -ForegroundColor Green
Write-Host "  BUILD CONCLUÍDO COM SUCESSO!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Arquivo gerado: SistemaPonto_Setup.exe ($setupSize MB)" -ForegroundColor White
Write-Host ""
Write-Host "  O instalador:" -ForegroundColor White
Write-Host "   ✓ Instala com duplo clique" -ForegroundColor Gray
Write-Host "   ✓ Cria atalho no Desktop e Menu Iniciar" -ForegroundColor Gray
Write-Host "   ✓ Abre sem terminal, sem Python visível" -ForegroundColor Gray
Write-Host "   ✓ Aparece no Painel de Controle para desinstalar" -ForegroundColor Gray
Write-Host ""
