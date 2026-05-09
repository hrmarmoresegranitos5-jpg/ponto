; ============================================================
;  NSIS — Script do Instalador  •  Sistema de Ponto v5.1
;  Gera: SistemaPonto_Setup.exe
;
;  Requisito: NSIS 3.x  (https://nsis.sourceforge.io)
;  Como usar:
;    1. Compilar o ponto.spec com PyInstaller → pasta dist\SistemaPonto\
;    2. Copiar o ícone icone.ico para esta pasta
;    3. Compilar este arquivo com makensis
; ============================================================

Unicode True

!include "MUI2.nsh"
!include "FileFunc.nsh"

; ── Informações do produto ──────────────────────────────────
!define APP_NAME        "Sistema de Ponto"
!define APP_VERSION     "5.1"
!define APP_PUBLISHER   "Sistema de Ponto"
!define APP_EXE         "SistemaPonto.exe"
!define APP_ICON        "icone.ico"
!define INSTALL_DIR     "$PROGRAMFILES64\SistemaPonto"
!define UNINSTALL_KEY   "Software\Microsoft\Windows\CurrentVersion\Uninstall\SistemaPonto"
!define DATA_DIR        "$APPDATA\SistemaPonto"

; ── Configurações do instalador ─────────────────────────────
Name              "${APP_NAME} ${APP_VERSION}"
OutFile           "SistemaPonto_Setup.exe"
InstallDir        "${INSTALL_DIR}"
InstallDirRegKey  HKLM "${UNINSTALL_KEY}" "InstallLocation"
RequestExecutionLevel admin
BrandingText      "${APP_NAME} v${APP_VERSION}"

; ── Aparência (MUI2) ────────────────────────────────────────
!define MUI_ICON                 "${APP_ICON}"
!define MUI_UNICON               "${APP_ICON}"
!define MUI_WELCOMEFINISHPAGE_BITMAP_NOSTRETCH
!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN       "$INSTDIR\${APP_EXE}"
!define MUI_FINISHPAGE_RUN_TEXT  "Abrir o Sistema de Ponto agora"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "PortugueseBR"

; ── Seção principal de instalação ──────────────────────────
Section "Principal" SecMain
    SectionIn RO   ; obrigatória, não pode desmarcar

    SetOutPath "$INSTDIR"
    SetOverwrite on

    ; Copia todos os arquivos da pasta dist\SistemaPonto\
    File /r "dist\SistemaPonto\*.*"

    ; Cria pasta de dados do usuário (onde fica o ponto.db)
    CreateDirectory "${DATA_DIR}"

    ; ── Atalho no Menu Iniciar ──────────────────────────────
    CreateDirectory "$SMPROGRAMS\${APP_NAME}"
    CreateShortcut  "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" \
                    "$INSTDIR\${APP_EXE}" "" \
                    "$INSTDIR\${APP_EXE}" 0 \
                    SW_SHOWNORMAL "" "${APP_NAME}"

    CreateShortcut  "$SMPROGRAMS\${APP_NAME}\Desinstalar ${APP_NAME}.lnk" \
                    "$INSTDIR\Uninstall.exe"

    ; ── Atalho na Área de Trabalho ─────────────────────────
    CreateShortcut  "$DESKTOP\${APP_NAME}.lnk" \
                    "$INSTDIR\${APP_EXE}" "" \
                    "$INSTDIR\${APP_EXE}" 0 \
                    SW_SHOWNORMAL "" "${APP_NAME}"

    ; ── Registro do desinstalador (Painel de Controle) ─────
    WriteRegStr   HKLM "${UNINSTALL_KEY}" "DisplayName"          "${APP_NAME}"
    WriteRegStr   HKLM "${UNINSTALL_KEY}" "DisplayVersion"       "${APP_VERSION}"
    WriteRegStr   HKLM "${UNINSTALL_KEY}" "Publisher"            "${APP_PUBLISHER}"
    WriteRegStr   HKLM "${UNINSTALL_KEY}" "InstallLocation"      "$INSTDIR"
    WriteRegStr   HKLM "${UNINSTALL_KEY}" "UninstallString"      "$INSTDIR\Uninstall.exe"
    WriteRegStr   HKLM "${UNINSTALL_KEY}" "DisplayIcon"          "$INSTDIR\${APP_EXE}"
    WriteRegDWORD HKLM "${UNINSTALL_KEY}" "NoModify"             1
    WriteRegDWORD HKLM "${UNINSTALL_KEY}" "NoRepair"             1

    ; Tamanho estimado instalado
    ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
    IntFmt $0 "0x%08X" $0
    WriteRegDWORD HKLM "${UNINSTALL_KEY}" "EstimatedSize" "$0"

    ; Escreve o desinstalador
    WriteUninstaller "$INSTDIR\Uninstall.exe"

SectionEnd

; ── Desinstalador ──────────────────────────────────────────
Section "Uninstall"

    ; Remove arquivos do programa
    RMDir /r "$INSTDIR"

    ; Remove atalhos
    Delete "$DESKTOP\${APP_NAME}.lnk"
    RMDir /r "$SMPROGRAMS\${APP_NAME}"

    ; Remove registro
    DeleteRegKey HKLM "${UNINSTALL_KEY}"

    ; Pergunta se deseja apagar os dados (ponto.db etc.)
    MessageBox MB_YESNO|MB_ICONQUESTION \
        "Deseja apagar os dados do programa ($APPDATA\SistemaPonto)?$\n$\nEscolha NÃO para manter o banco de dados." \
        IDNO keep_data

    RMDir /r "${DATA_DIR}"
    keep_data:

SectionEnd
