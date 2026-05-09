"""
╔══════════════════════════════════════════════════════════════════╗
║  SERVIDOR WEB — BANCO DE HORAS  (versão nuvem)                  ║
║  Suporta SQLite (local) e PostgreSQL (Railway/Render/nuvem)     ║
╚══════════════════════════════════════════════════════════════════╝

Variáveis de ambiente necessárias na nuvem:
    DATABASE_URL  → URL do PostgreSQL (ex: postgresql://user:pass@host/db)
                    Se não definida, usa SQLite local (ponto.db)
    SECRET_KEY    → Chave secreta opcional para segurança extra

Como fazer deploy (Railway):
    1. Suba o código para o GitHub
    2. Conecte o repositório no Railway
    3. Adicione um banco PostgreSQL no Railway
    4. A variável DATABASE_URL é preenchida automaticamente
    5. Deploy automático!
"""

import os
import secrets
import socket
from datetime import datetime, timedelta
from functools import wraps

try:
    from flask import Flask, request, jsonify, send_from_directory
except ImportError:
    print("\n" + "="*60)
    print("  ERRO: Flask não instalado.")
    print("  Execute:  pip install flask")
    print("="*60 + "\n")
    raise

# ──────────────────────────────────────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────────────────────────────────────
DATABASE_URL = os.environ.get("DATABASE_URL", "")
DB_FILE      = os.path.join(os.path.dirname(__file__) or ".", "ponto.db")
STATIC_DIR   = os.path.dirname(__file__) or "."
PORT         = int(os.environ.get("PORT", 5000))
TOKEN_TTL    = timedelta(hours=8)

# Railway usa postgres://, psycopg2 precisa de postgresql://
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

USE_POSTGRES = bool(DATABASE_URL)

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path="/static")
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))

# Tokens em memória: token → { func_id, nome, expires }
_tokens: dict = {}


# ──────────────────────────────────────────────────────────────────────────────
# CAMADA DE BANCO DE DADOS  (SQLite ou PostgreSQL transparente)
# ──────────────────────────────────────────────────────────────────────────────
def _conn():
    if USE_POSTGRES:
        import psycopg2
        import psycopg2.extras
        return psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    else:
        import sqlite3
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        return conn


def _ph():
    """Placeholder: ? para SQLite, %s para PostgreSQL."""
    return "%s" if USE_POSTGRES else "?"


def _run(conn, sql, params=()):
    """Executa SQL trocando ? pelo placeholder correto."""
    cur = conn.cursor()
    cur.execute(sql.replace("?", _ph()), params)
    return cur


def _criar_tabelas():
    conn = _conn()
    cur = conn.cursor()
    if USE_POSTGRES:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS funcionarios (
                id SERIAL PRIMARY KEY,
                nome TEXT NOT NULL UNIQUE,
                salario REAL NOT NULL DEFAULT 0,
                tipo_pag TEXT NOT NULL DEFAULT 'mes',
                jornada_h REAL NOT NULL DEFAULT 8,
                observacoes TEXT,
                ativo INTEGER NOT NULL DEFAULT 1,
                criado_em TEXT NOT NULL DEFAULT '',
                pin TEXT DEFAULT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS banco_horas (
                id SERIAL PRIMARY KEY,
                funcionario_id INTEGER NOT NULL UNIQUE,
                saldo_min INTEGER NOT NULL DEFAULT 0,
                atualizado_em TEXT NOT NULL DEFAULT ''
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS banco_horas_mov (
                id SERIAL PRIMARY KEY,
                funcionario_id INTEGER NOT NULL,
                tipo TEXT NOT NULL,
                minutos INTEGER NOT NULL,
                descricao TEXT,
                referencia TEXT,
                criado_em TEXT NOT NULL DEFAULT ''
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS importacoes (
                id SERIAL PRIMARY KEY,
                arquivo TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'ok',
                importado_em TEXT NOT NULL DEFAULT '',
                total_reg INTEGER DEFAULT 0,
                total_erros INTEGER DEFAULT 0,
                avisos TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS batidas (
                id SERIAL PRIMARY KEY,
                importacao_id INTEGER NOT NULL,
                funcionario_id INTEGER,
                funcionario_raw TEXT NOT NULL,
                data TEXT NOT NULL,
                horarios TEXT NOT NULL,
                qtd_batidas INTEGER NOT NULL DEFAULT 0,
                tem_erro INTEGER NOT NULL DEFAULT 0,
                erros TEXT
            )
        """)
    else:
        cur.executescript("""
            CREATE TABLE IF NOT EXISTS funcionarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL UNIQUE, salario REAL NOT NULL DEFAULT 0,
                tipo_pag TEXT NOT NULL DEFAULT 'mes', jornada_h REAL NOT NULL DEFAULT 8,
                observacoes TEXT, ativo INTEGER NOT NULL DEFAULT 1,
                criado_em TEXT NOT NULL DEFAULT '', pin TEXT DEFAULT NULL
            );
            CREATE TABLE IF NOT EXISTS banco_horas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                funcionario_id INTEGER NOT NULL UNIQUE,
                saldo_min INTEGER NOT NULL DEFAULT 0,
                atualizado_em TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS banco_horas_mov (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                funcionario_id INTEGER NOT NULL, tipo TEXT NOT NULL,
                minutos INTEGER NOT NULL, descricao TEXT,
                referencia TEXT, criado_em TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS importacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                arquivo TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'ok',
                importado_em TEXT NOT NULL DEFAULT '', total_reg INTEGER DEFAULT 0,
                total_erros INTEGER DEFAULT 0, avisos TEXT
            );
            CREATE TABLE IF NOT EXISTS batidas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                importacao_id INTEGER NOT NULL, funcionario_id INTEGER,
                funcionario_raw TEXT NOT NULL, data TEXT NOT NULL,
                horarios TEXT NOT NULL, qtd_batidas INTEGER NOT NULL DEFAULT 0,
                tem_erro INTEGER NOT NULL DEFAULT 0, erros TEXT
            );
        """)
    conn.commit()
    conn.close()


# ──────────────────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────────────────
def _fmt(mins: int) -> str:
    m = abs(mins)
    return f"{m // 60:02d}:{m % 60:02d}"


def _saldo_fmt(mins: int) -> str:
    return f"{'-' if mins < 0 else '+'}{_fmt(mins)}"


def _d(row):
    """Converte Row (sqlite3 ou psycopg2 RealDictRow) para dict."""
    if row is None:
        return None
    return dict(row)


# ──────────────────────────────────────────────────────────────────────────────
# CONSULTAS
# ──────────────────────────────────────────────────────────────────────────────
def buscar_func_por_nome(nome: str):
    conn = _conn()
    cur  = _run(conn, "SELECT id, nome FROM funcionarios WHERE nome=? AND ativo=1", (nome,))
    row  = cur.fetchone()
    conn.close()
    return _d(row)


def buscar_func_por_pin(pin: str):
    conn = _conn()
    cur  = _run(conn, "SELECT id, nome FROM funcionarios WHERE pin=? AND ativo=1", (pin,))
    row  = cur.fetchone()
    conn.close()
    return _d(row)


def buscar_saldo(func_id: int) -> dict:
    conn = _conn()

    cur = _run(conn,
        "SELECT saldo_min, atualizado_em FROM banco_horas WHERE funcionario_id=?",
        (func_id,))
    saldo_row = _d(cur.fetchone())

    cur = _run(conn, """
        SELECT tipo, minutos, descricao, referencia, criado_em
        FROM banco_horas_mov
        WHERE funcionario_id=?
        ORDER BY criado_em DESC
        LIMIT 20
    """, (func_id,))
    movs = [_d(r) for r in cur.fetchall()]

    cur = _run(conn, "SELECT nome FROM funcionarios WHERE id=?", (func_id,))
    func = _d(cur.fetchone())

    conn.close()

    saldo_min  = saldo_row["saldo_min"]    if saldo_row else 0
    atualizado = saldo_row["atualizado_em"] if saldo_row else None
    nome       = func["nome"] if func else "?"

    historico = []
    for m in movs:
        is_cred = m["tipo"] == "credito"
        historico.append({
            "tipo":      m["tipo"],
            "minutos":   m["minutos"],
            "valor_fmt": ("+" if is_cred else "−") + _fmt(m["minutos"]),
            "descricao": m["descricao"] or "",
            "data":      (m["criado_em"] or "")[:16],
        })

    return {
        "nome":           nome,
        "saldo_min":      saldo_min,
        "saldo_fmt":      _saldo_fmt(saldo_min),
        "saldo_horas":    abs(saldo_min) // 60,
        "saldo_mins":     abs(saldo_min) % 60,
        "saldo_positivo": saldo_min >= 0,
        "atualizado_em":  (atualizado or "")[:16],
        "historico":      historico,
    }


def listar_funcionarios_resumo() -> list:
    conn = _conn()
    cur  = _run(conn, """
        SELECT f.id, f.nome, COALESCE(b.saldo_min, 0) as saldo_min
        FROM funcionarios f
        LEFT JOIN banco_horas b ON b.funcionario_id = f.id
        WHERE f.ativo = 1
        ORDER BY f.nome
    """)
    rows = [_d(r) for r in cur.fetchall()]
    conn.close()
    return [
        {
            "id":       r["id"],
            "nome":     r["nome"],
            "saldo":    _saldo_fmt(r["saldo_min"]),
            "positivo": r["saldo_min"] >= 0,
        }
        for r in rows
    ]


# ──────────────────────────────────────────────────────────────────────────────
# AUTH
# ──────────────────────────────────────────────────────────────────────────────
def _limpar_tokens_expirados():
    agora     = datetime.now()
    expirados = [t for t, d in _tokens.items() if d["expires"] < agora]
    for t in expirados:
        del _tokens[t]


def _token_valido(token: str):
    _limpar_tokens_expirados()
    return _tokens.get(token)


def requer_token(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.headers.get("X-Token", "")
        sess  = _token_valido(token)
        if not sess:
            return jsonify({"erro": "Não autorizado. Faça login novamente."}), 401
        request.func_id   = sess["func_id"]
        request.func_nome = sess["nome"]
        return f(*args, **kwargs)
    return decorated


# ──────────────────────────────────────────────────────────────────────────────
# ROTAS DA API
# ──────────────────────────────────────────────────────────────────────────────
@app.route("/api/login", methods=["POST"])
def api_login():
    dados = request.get_json(silent=True) or {}
    pin   = str(dados.get("pin", "")).strip()
    if not pin:
        return jsonify({"erro": "PIN não informado."}), 400
    func = buscar_func_por_pin(pin)
    if not func:
        return jsonify({"erro": "PIN inválido."}), 401
    token = secrets.token_hex(24)
    _tokens[token] = {
        "func_id": func["id"],
        "nome":    func["nome"],
        "expires": datetime.now() + TOKEN_TTL,
    }
    return jsonify({"token": token, "nome": func["nome"]}), 200


@app.route("/api/logout", methods=["POST"])
@requer_token
def api_logout():
    token = request.headers.get("X-Token", "")
    _tokens.pop(token, None)
    return jsonify({"ok": True}), 200


@app.route("/api/saldo")
@requer_token
def api_saldo():
    return jsonify(buscar_saldo(request.func_id)), 200


@app.route("/api/funcionarios")
def api_funcionarios():
    """Lista resumida de todos os funcionários com saldo."""
    return jsonify(listar_funcionarios_resumo()), 200


@app.route("/api/saldo-por-nome/<nome>")
def api_saldo_por_nome(nome):
    """
    Retorna o saldo de um funcionário pelo nome.
    Não exige token — a segurança vem do app ficar bloqueado no dispositivo.
    """
    func = buscar_func_por_nome(nome)
    if not func:
        return jsonify({"erro": "Funcionário não encontrado."}), 404
    return jsonify(buscar_saldo(func["id"])), 200


@app.route("/api/health")
def api_health():
    """Healthcheck — Railway verifica se o serviço está no ar."""
    return jsonify({
        "status": "ok",
        "banco":  "postgresql" if USE_POSTGRES else "sqlite",
        "hora":   datetime.now().isoformat(),
    }), 200


# ──────────────────────────────────────────────────────────────────────────────
# SERVE O APP (PWA)
# ──────────────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.route("/manifest.json")
def manifest():
    return send_from_directory(STATIC_DIR, "manifest.json")


@app.route("/sw.js")
def service_worker():
    resp = send_from_directory(STATIC_DIR, "sw.js")
    resp.headers["Cache-Control"] = "no-cache"
    resp.headers["Service-Worker-Allowed"] = "/"
    return resp


# ──────────────────────────────────────────────────────────────────────────────
# INICIALIZAÇÃO
# ──────────────────────────────────────────────────────────────────────────────
def _ip_local():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


# Cria as tabelas ao iniciar (seguro rodar várias vezes)
with app.app_context():
    try:
        _criar_tabelas()
        print(f"[OK] Banco inicializado ({'PostgreSQL' if USE_POSTGRES else 'SQLite'})")
    except Exception as e:
        print(f"[AVISO] Erro ao inicializar banco: {e}")


if __name__ == "__main__":
    modo = "PostgreSQL" if USE_POSTGRES else f"SQLite ({DB_FILE})"
    ip   = _ip_local()
    print("\n" + "="*60)
    print("  SERVIDOR BANCO DE HORAS — rodando!")
    print(f"  Banco de dados:  {modo}")
    print("="*60)
    print(f"  Acesso local:    http://localhost:{PORT}")
    print(f"  Acesso na rede:  http://{ip}:{PORT}")
    print("="*60 + "\n")
    app.run(host="0.0.0.0", port=PORT, debug=False)
