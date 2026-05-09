# 🖥️ SISTEMA DE PONTO — GUIA DE INSTALAÇÃO NO PC

---

## ✅ PRÉ-REQUISITOS

Antes de começar, verifique se o seu PC tem:
- **Windows 10 ou 11**
- **Python 3.10 ou superior** → https://python.org/downloads
  - Na hora de instalar marque: ☑ "Add Python to PATH"

---

## 📁 PASSO 1 — Colocar os arquivos no PC

1. Extraia o ZIP em uma pasta de fácil acesso, por exemplo:
   ```
   C:\SistemaPonto\
   ```
2. A pasta deve ficar assim:
   ```
   C:\SistemaPonto\
     ├── ponto.py          ← Programa principal
     ├── servidor.py       ← Servidor da nuvem (não roda no PC)
     ├── relatorios_pdf.py ← Gerador de relatórios
     ├── index.html        ← App do celular
     └── ...
   ```

---

## 📦 PASSO 2 — Instalar as bibliotecas necessárias

1. Abra o **Prompt de Comando** (tecla `Win` → digite `cmd` → Enter)
2. Digite os comandos abaixo um por um:

```cmd
pip install openpyxl xlrd reportlab requests
```

Aguarde terminar. Se aparecer erros, tente:
```cmd
python -m pip install --upgrade pip
python -m pip install openpyxl xlrd reportlab requests
```

---

## ▶️ PASSO 3 — Rodar o programa

No Prompt de Comando, navegue até a pasta e execute:

```cmd
cd C:\SistemaPonto
python ponto.py
```

Ou crie um atalho: clique com botão direito em `ponto.py` →
**"Abrir com" → Python** (selecione "sempre usar esse programa").

---

## ☁️ PASSO 4 — Conectar ao servidor da nuvem (Railway)

> Esse passo faz os dados aparecerem no celular dos funcionários.

### 4.1 — Subir o código no GitHub
Se ainda não fez:
```cmd
git init
git add .
git commit -m "Sistema de ponto v5"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/sistema-ponto.git
git push -u origin main
```

### 4.2 — Configurar o Railway
1. Acesse **railway.app** e faça login
2. Vá no seu projeto → **Variables** → adicione:
   ```
   ADMIN_KEY = uma_senha_sua_qualquer_ex_ponto2025
   ```
3. Faça um novo deploy (push no GitHub ou botão "Deploy" no Railway)

### 4.3 — Configurar a URL no programa do PC
1. Abra o `ponto.py` normalmente
2. No menu lateral, clique em **⚙ Configurações**
3. Na seção **"☁ Sincronização com a Nuvem"**:
   - **URL do Servidor:** `https://SEU-PROJETO.up.railway.app`
   - **ADMIN_KEY:** a mesma senha que você colocou no Railway
4. Clique em **💾 Salvar URL / Chave**
5. Clique em **↑ Sincronizar Agora** para testar

Se aparecer "✔ Sincronizado!" está tudo funcionando! 🎉

---

## 📱 PASSO 5 — App dos funcionários no celular

1. Abra a URL do Railway no celular:
   ```
   https://SEU-PROJETO.up.railway.app
   ```
2. O navegador vai perguntar se quer "Adicionar à tela inicial" → aceite
3. O funcionário digita o PIN e consulta o saldo

| Funcionário | PIN  |
|-------------|------|
| Fabricio    | 1234 |
| Gibs        | 2345 |
| Hugo        | 3456 |
| Tiago       | 4567 |

> ⚠️ Troque os PINs antes de usar!
> Menu: Banco de Horas → selecionar funcionário → campo PIN → Salvar PIN

---

## 🔄 COMO FUNCIONA A SINCRONIZAÇÃO

| O que você faz no PC | O que acontece no celular |
|---|---|
| Importar arquivo de ponto | Saldo atualizado automaticamente |
| Lançar retirada no banco de horas | Funcionário já vê o novo saldo |
| Clicar "Sincronizar Agora" | Força atualização de todos |

---

## ❓ PROBLEMAS COMUNS

### "python não é reconhecido"
→ Python não foi instalado com o PATH. Reinstale e marque ☑ "Add Python to PATH".

### "ModuleNotFoundError: No module named 'openpyxl'"
→ Execute novamente: `pip install openpyxl xlrd reportlab requests`

### "Falha na conexão" ao sincronizar
→ Verifique se a URL está correta e se o Railway está rodando.
→ Confira se a ADMIN_KEY no Railway é igual à do programa.

### O programa abre e fecha rápido
→ Abra pelo Prompt de Comando (`python ponto.py`) para ver a mensagem de erro.

---

## 📂 ARQUIVOS GERADOS AUTOMATICAMENTE

Após a primeira execução, serão criados na mesma pasta:
- `ponto.db` → banco de dados local (não apagar!)
- `ponto_config.ini` → configurações de nuvem (URL e chave)

---

Dúvidas? Consulte o `README.md` para detalhes técnicos completos.
