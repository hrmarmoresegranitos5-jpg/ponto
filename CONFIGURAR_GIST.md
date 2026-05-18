# 🚀 Como Configurar o Sistema de Ponto (GitHub Gist)

Sem servidor, sem Railway. Tudo funciona pelo GitHub gratuito.

---

## Arquitetura nova

```
ponto.py  →  GitHub Gist (JSON público)  ←  index.html (PWA no celular)
  (PC)            (backend gratuito)           (funcionários)
```

---

## Passo 1 — Gerar o Token do GitHub

1. Acesse: **https://github.com/settings/tokens**
2. Clique em **"Generate new token (classic)"**
3. Nome: `Sistema de Ponto`
4. Selecione somente o escopo **`gist`**
5. Clique em **Generate token**
6. **Copie o token** (só aparece uma vez!)

---

## Passo 2 — Configurar no ponto.py

1. Abra o `ponto.py`
2. Vá em **Configurações**
3. Na seção **GitHub Gist**:
   - Cole o token no campo **GitHub Token (PAT)**
   - Deixe o **ID do Gist** vazio (será criado automaticamente)
4. Clique em **Salvar Token / ID**

---

## Passo 3 — Primeira Sincronização

1. Ainda em Configurações, clique **Sincronizar Agora**
2. Uma janela aparecerá com o **ID do Gist criado** e a **URL completa**
3. **Copie a URL** — ela vai parecer com:
   ```
   https://gist.githubusercontent.com/SEU_USUARIO/abc123.../raw/data.json
   ```

---

## Passo 4 — Hospedar o index.html no GitHub Pages

1. Crie um repositório no GitHub (ex: `sistema-ponto`)
2. Faça upload dos arquivos: `index.html`, `manifest.json`, `sw.js`
3. Vá em **Settings → Pages → Source: main branch / root**
4. Sua URL será: `https://SEU_USUARIO.github.io/sistema-ponto/`

---

## Passo 5 — Colocar a URL no index.html

Abra o `index.html` e localize (linha ~2200):

```javascript
const GIST_RAW_URL = '';
```

Substitua por:

```javascript
const GIST_RAW_URL = 'https://gist.githubusercontent.com/SEU_USUARIO/ID/raw/data.json';
```

Salve e faça commit no GitHub Pages.

---

## Uso do dia a dia

1. Importe o arquivo do relógio de ponto no `ponto.py` normalmente
2. Os dados são **sincronizados automaticamente** com o Gist após cada importação
3. Os funcionários abrem o app no celular → dados atualizados em até 5 minutos

---

## ✅ O que funciona offline no celular

- Saldo acumulado (cache local de 5 min)
- Histórico de batidas
- Notificações já recebidas

## ⚠️ Não precisa mais

- ❌ Railway
- ❌ PostgreSQL
- ❌ Heroku
- ❌ Qualquer servidor
