# Sistema de Ponto — v5.0

App para leitura de relatórios de relógio de ponto e controle de banco de horas.  
Os funcionários consultam o saldo pelo celular via PWA (sem instalar da loja).

---

## Arquitetura

```
ponto.py          → Roda no PC da empresa (Windows). Importa os arquivos
                    do relógio e calcula as horas extras.
servidor.py       → API Flask. Roda na nuvem (Railway). Funcionários
                    consultam pelo celular de qualquer lugar.
index.html        → App do celular (PWA).
relatorios_pdf.py → Geração de PDF e Excel (usado pelo ponto.py).
ponto.db          → Banco SQLite local (gerado pelo ponto.py).
```

---

## ▶ COMO FAZER O DEPLOY NA NUVEM (Railway)

### Passo 1 — Subir o código no GitHub

```bash
git init
git add .
git commit -m "Sistema de ponto v5"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/sistema-ponto.git
git push -u origin main
```

### Passo 2 — Criar conta e projeto no Railway

1. Acesse **[railway.app](https://railway.app)** e crie uma conta gratuita
2. Clique em **New Project → Deploy from GitHub repo**
3. Selecione o repositório `sistema-ponto`
4. O Railway detecta o `Procfile` e inicia o deploy automaticamente

### Passo 3 — Adicionar o banco de dados PostgreSQL

1. No projeto Railway, clique em **+ New → Database → PostgreSQL**
2. O Railway cria o banco e define a variável `DATABASE_URL` automaticamente
3. **Não é necessário configurar nada** — o `servidor.py` detecta e usa

### Passo 4 — Obter a URL pública

1. No serviço do Railway, vá em **Settings → Networking → Generate Domain**
2. Você terá uma URL como `https://sistema-ponto-production.up.railway.app`
3. Compartilhe essa URL com os funcionários

### Passo 5 — Conectar o ponto.py ao banco da nuvem

No `ponto.py`, adicione a URL do PostgreSQL como variável de ambiente:

**Windows (Prompt de Comando):**
```cmd
set DATABASE_URL=postgresql://usuario:senha@host/db
python ponto.py
```

**Ou crie um arquivo `.env` na pasta do ponto.py:**
```
DATABASE_URL=postgresql://usuario:senha@host/db
```
*(A string de conexão fica disponível em Railway → PostgreSQL → Connect)*

---

## ▶ COMO USAR LOCALMENTE (sem nuvem)

### Requisitos

```bash
pip install flask psycopg2-binary gunicorn openpyxl xlrd reportlab
```

### Programa desktop (Windows)

```bash
python ponto.py
```

### Servidor web (celular na mesma rede)

```bash
python servidor.py
```
Acesse o endereço mostrado no terminal pelo celular na mesma Wi-Fi.

---

## Funcionários pré-cadastrados

| Funcionário | PIN  |
|-------------|------|
| Fabricio    | 1234 |
| Gibs        | 2345 |
| Hugo        | 3456 |
| Tiago       | 4567 |

> **Troque os PINs** antes de colocar em produção!  
> Menu: Banco de Horas → selecionar funcionário → campo PIN → Salvar PIN

---

## Formatos de arquivo aceitos pelo relógio de ponto

| Formato | Exemplo |
|---------|---------|
| CSV com nome | `Funcionario,Data,Horario` |
| TXT tipo REP | Matrícula / Data / Hora por linha |
| Excel `.xlsx` | Mesma estrutura do CSV |

---

## Rotas da API

| Método | Rota | Descrição |
|--------|------|-----------|
| `POST` | `/api/login` | `{ "pin": "1234" }` → retorna token |
| `GET`  | `/api/saldo` | Saldo do funcionário logado |
| `POST` | `/api/logout` | Invalida o token |
| `GET`  | `/api/funcionarios` | Lista todos os funcionários com saldo |
| `GET`  | `/api/health` | Verifica se o servidor está no ar |

---

## Etapas do projeto

| Etapa | Status | Descrição |
|-------|--------|-----------|
| 1 | ✅ | Estrutura, banco de dados, cadastro de funcionários |
| 2 | ✅ | Importação CSV / TXT / XLS / XLSX |
| 3 | ✅ | Cálculo de horas extras, faltas e saldo |
| 4 | ✅ | Banco de horas acumulado + retiradas + app celular |
| 5 | ✅ | Relatórios PDF e exportação Excel |
| 6 | ✅ | Deploy na nuvem (Railway + PostgreSQL) |
| 7 | —  | Backup automático |
