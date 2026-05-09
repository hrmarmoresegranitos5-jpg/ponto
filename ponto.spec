# -*- mode: python ; coding: utf-8 -*-
# ============================================================
#  PyInstaller SPEC — Sistema de Ponto v5.1
#  Gera: dist/SistemaPonto/SistemaPonto.exe  (one-folder)
# ============================================================

import sys
from pathlib import Path

block_cipher = None

# Arquivos extras que devem ir junto com o executável
added_files = [
    # (origem,  destino_dentro_do_pacote)
    ('relatorios_pdf.py', '.'),
    ('notificacoes.py',   '.'),
    ('index.html',        '.'),
    ('manifest.json',     '.'),
    ('sw.js',             '.'),
]

a = Analysis(
    ['ponto.py'],
    pathex=['.'],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'sqlite3',
        'openpyxl',
        'xlrd',
        'reportlab',
        'reportlab.pdfgen',
        'reportlab.lib',
        'reportlab.lib.pagesizes',
        'reportlab.lib.styles',
        'reportlab.lib.units',
        'reportlab.platypus',
        'requests',
        'configparser',
        'relatorios_pdf',
        'notificacoes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'servidor.py',   # roda só na nuvem
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
        'PyQt5',
        'PyQt6',
        'wx',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SistemaPonto',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # <-- SEM janela de terminal!
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icone.ico',        # ícone personalizado
    version='version_info.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SistemaPonto',
)
