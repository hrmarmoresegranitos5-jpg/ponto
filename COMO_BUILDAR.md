# 📦 Sistema de Ponto — Criando o Instalador Windows (.exe)

Este guia transforma o `ponto.py` num instalador profissional para Windows.

---

## O que você vai ter no final

```
SistemaPonto_Setup.exe   ← instalador com duplo clique
```

Ao rodar o instalador:
- ✅ Instala o programa em `C:\Program Files\SistemaPonto\`
- ✅ Cria atalho na **Área de Trabalho**
- ✅ Cria atalho no **Menu Iniciar**
- ✅ Abre **sem terminal**, **sem Python visível**
- ✅ Aparece no **Painel de Controle → Programas** para desinstalar

---

## Pré-requisitos (instalar uma única vez)

### 1. Python 3.10+
- Baixe em: https://python.org/downloads
- ⚠️ **IMPORTANTE:** na instalação, marque ☑ **"Add Python to PATH"**

### 2. NSIS (criador do instalador)
- Baixe em: https://nsis.sourceforge.io/Download
- Instale normalmente (Next → Next → Install)
- Após instalar, adicione ao PATH:
  - `C:\Program Files (x86)\NSIS\` ou `C:\Program Files\NSIS\`

---

## Como buildar (passo a passo)

### Opção A — Automático (recomendado)

1. Coloque **todos os arquivos** desta pasta numa mesma pasta no seu PC, por exemplo:
   ```
   C:\BuildPonto\
   ```

2. Abra o **PowerShell** nessa pasta (Shift + clique direito → "Abrir PowerShell aqui")

3. Execute:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\BUILD.ps1
   ```

4. Aguarde ~2 minutos. O arquivo `SistemaPonto_Setup.exe` será criado na mesma pasta.

> Se preferir usar o CMD, execute `BUILD.bat` ao invés do BUILD.ps1.

---

### Opção B — Manual (passo a passo)

**Passo 1 — Instalar dependências:**
```cmd
pip install pyinstaller openpyxl xlrd reportlab requests
```

**Passo 2 — Gerar o ícone:**
```cmd
python gerar_icone.py
```

**Passo 3 — Compilar o executável:**
```cmd
pyinstaller ponto.spec --noconfirm
```
Isso cria a pasta `dist\SistemaPonto\` com o `SistemaPonto.exe` dentro.

**Passo 4 — Criar o instalador:**
```cmd
makensis installer.nsi
```
Isso cria o `SistemaPonto_Setup.exe`.

---

## Estrutura dos arquivos

```
C:\BuildPonto\
  ├── ponto.py              ← Programa principal (não modificar)
  ├── relatorios_pdf.py     ← Módulo de relatórios
  ├── notificacoes.py       ← Módulo de notificações
  ├── index.html            ← App mobile (vai junto no instalador)
  ├── manifest.json         ← PWA manifest
  ├── sw.js                 ← Service Worker
  │
  ├── ponto.spec            ← Config do PyInstaller (não modificar)
  ├── version_info.txt      ← Metadados do .exe
  ├── installer.nsi         ← Script do instalador NSIS
  ├── gerar_icone.py        ← Gerador de ícone automático
  │
  ├── BUILD.ps1             ← Script de build (PowerShell)
  └── BUILD.bat             ← Script de build (CMD)
```

Após o build, serão criados:
```
  ├── icone.ico             ← Ícone gerado automaticamente
  ├── build\                ← Temporário (pode apagar)
  ├── dist\SistemaPonto\    ← Executável e DLLs
  └── SistemaPonto_Setup.exe ← ← ← ESTE É O INSTALADOR FINAL
```

---

## Personalizações

### Trocar o ícone
Substitua o `icone.ico` por um arquivo `.ico` da sua preferência antes do build.
Ferramentas online para criar ICO: https://favicon.io ou https://icoconvert.com

### Trocar o nome/versão
Edite as primeiras linhas do `installer.nsi`:
```nsi
!define APP_NAME    "Minha Empresa — Ponto"
!define APP_VERSION "6.0"
```

### Alterar pasta de instalação padrão
No `installer.nsi`, altere:
```nsi
!define INSTALL_DIR "$PROGRAMFILES64\SistemaPonto"
```

---

## Perguntas frequentes

**Q: O antivírus acusou o Setup.exe como vírus.**
A: Isso é comum em executáveis gerados com PyInstaller — é um falso positivo.
   O PyInstaller "embala" Python num único .exe, o que parece suspeito para alguns AVs.
   Soluções: assinar o executável digitalmente (certificate signing) ou adicionar exceção no antivírus.

**Q: O programa abre e fecha imediatamente após instalar.**
A: Abra o CMD e rode direto:
   ```cmd
   "C:\Program Files\SistemaPonto\SistemaPonto.exe"
   ```
   O erro aparecerá no terminal.

**Q: Posso distribuir para vários PCs?**
A: Sim. Copie o `SistemaPonto_Setup.exe` para um pendrive ou compartilhamento de rede
   e execute em cada PC. Não precisa de Python instalado nos PCs de destino.

**Q: O banco de dados (ponto.db) fica onde?**
A: Na mesma pasta do executável: `C:\Program Files\SistemaPonto\ponto.db`
   Ao desinstalar, o instalador pergunta se deseja apagar ou manter os dados.

---

## Tamanho esperado

| Arquivo           | Tamanho aprox. |
|-------------------|---------------|
| dist\SistemaPonto\ (pasta) | 30–60 MB |
| SistemaPonto_Setup.exe     | 15–35 MB |

O tamanho varia dependendo das versões das bibliotecas.

---

*Gerado automaticamente para Sistema de Ponto v5.1*
