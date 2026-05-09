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
ADMIN_KEY    = os.environ.get("ADMIN_KEY", "")  # Chave para sincronização do ponto.py

# Railway usa postgres://, psycopg2 precisa de postgresql://
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

USE_POSTGRES = bool(DATABASE_URL)

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path="/static")
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))


@app.after_request
def _cors(response):
    """Permite que o GitHub Pages (ou qualquer origem) acesse a API."""
    response.headers["Access-Control-Allow-Origin"]  = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Token"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return response


@app.route("/api/<path:_>", methods=["OPTIONS"])
def _options(_):
    return "", 204

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


def _criar_tabelas_multiplicadores(conn):
    """Cria tabelas de configuração de multiplicadores de hora extra."""
    # Multiplicador global (padrão do sistema)
    _run(conn, """
        CREATE TABLE IF NOT EXISTS mult_config_global (
            id          INTEGER PRIMARY KEY DEFAULT 1,
            multiplicador REAL NOT NULL DEFAULT 1.5,
            atualizado_em TEXT NOT NULL
        )
    """)
    # Multiplicador por funcionário (sobrepõe o global)
    _run(conn, """
        CREATE TABLE IF NOT EXISTS mult_config_funcionario (
            funcionario_id INTEGER PRIMARY KEY,
            multiplicador  REAL NOT NULL DEFAULT 1.5,
            atualizado_em  TEXT NOT NULL
        )
    """)
    # Multiplicador por dia da semana (0=Dom..6=Sáb; sobrepõe funcionário e global)
    _run(conn, """
        CREATE TABLE IF NOT EXISTS mult_config_dia_semana (
            dia_semana    INTEGER NOT NULL,
            multiplicador REAL NOT NULL DEFAULT 1.5,
            atualizado_em TEXT NOT NULL,
            PRIMARY KEY (dia_semana)
        )
    """)
    # Multiplicador por período com datas (maior prioridade)
    _run(conn, """
        CREATE TABLE IF NOT EXISTS mult_config_periodo (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            funcionario_id INTEGER,
            data_inicio    TEXT NOT NULL,
            data_fim       TEXT NOT NULL,
            multiplicador  REAL NOT NULL DEFAULT 1.5,
            descricao      TEXT,
            criado_em      TEXT NOT NULL
        )
    """)
    # Auditoria: multiplicador usado em cada fechamento semanal
    _run(conn, """
        CREATE TABLE IF NOT EXISTS mult_auditoria (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            funcionario_id INTEGER NOT NULL,
            referencia     TEXT NOT NULL,
            multiplicador  REAL NOT NULL,
            valor_hora     REAL NOT NULL,
            valor_hora_ext REAL NOT NULL,
            saldo_min      INTEGER NOT NULL,
            valor_total    REAL NOT NULL,
            criado_em      TEXT NOT NULL,
            observacao     TEXT
        )
    """)
    conn.commit()


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
        cur.execute("""
            CREATE TABLE IF NOT EXISTS pagamentos_horas (
                id SERIAL PRIMARY KEY,
                funcionario_id INTEGER NOT NULL,
                minutos_pagos INTEGER NOT NULL,
                valor_pago REAL NOT NULL,
                tipo TEXT NOT NULL,
                descricao TEXT,
                criado_em TEXT NOT NULL
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
            CREATE TABLE IF NOT EXISTS pagamentos_horas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                funcionario_id INTEGER NOT NULL,
                minutos_pagos INTEGER NOT NULL,
                valor_pago REAL NOT NULL,
                tipo TEXT NOT NULL,
                descricao TEXT,
                criado_em TEXT NOT NULL
            );
        """)
    conn.commit()
    # Cria tabelas de multiplicadores (sempre seguro chamar)
    _criar_tabelas_multiplicadores(conn)
    conn.close()


# ──────────────────────────────────────────────────────────────────────────────
# HELPERS DE MULTIPLICADOR
# ──────────────────────────────────────────────────────────────────────────────
def _get_multiplicador(func_id: int, data_str: str = None) -> float:
    """
    Retorna o multiplicador efetivo para um funcionário em uma data,
    seguindo a hierarquia de prioridade:
      1. Período específico (maior prioridade)
      2. Dia da semana
      3. Por funcionário
      4. Global (menor prioridade)
    """
    conn = _conn()
    try:
        _criar_tabelas_multiplicadores(conn)
    except Exception:
        pass

    mult = 1.5  # fallback padrão

    # 4. Global
    try:
        row = _d(_run(conn, "SELECT multiplicador FROM mult_config_global WHERE id=1").fetchone())
        if row:
            mult = row["multiplicador"]
    except Exception:
        pass

    # 3. Por funcionário
    try:
        row = _d(_run(conn,
            "SELECT multiplicador FROM mult_config_funcionario WHERE funcionario_id=?",
            (func_id,)).fetchone())
        if row:
            mult = row["multiplicador"]
    except Exception:
        pass

    # 2. Dia da semana (se data fornecida)
    if data_str:
        try:
            from datetime import date as ddate
            if "-" in data_str:
                ano, mes, dia = data_str.split("-")[:3]
            else:
                dia, mes, ano = data_str.split("/")[:3]
            dt = ddate(int(ano), int(mes), int(dia))
            dow = dt.weekday()  # 0=Seg..6=Dom → converte para 0=Dom..6=Sáb
            dow_sun = (dow + 1) % 7
            row = _d(_run(conn,
                "SELECT multiplicador FROM mult_config_dia_semana WHERE dia_semana=?",
                (dow_sun,)).fetchone())
            if row:
                mult = row["multiplicador"]
        except Exception:
            pass

    # 1. Período específico (maior prioridade)
    if data_str:
        try:
            if "-" in data_str:
                data_norm = data_str[:10]
            else:
                d, m, a = data_str.split("/")[:3]
                data_norm = f"{a}-{m}-{d}"
            row = _d(_run(conn, """
                SELECT multiplicador FROM mult_config_periodo
                WHERE (funcionario_id IS NULL OR funcionario_id=?)
                  AND data_inicio <= ? AND data_fim >= ?
                ORDER BY funcionario_id DESC, id DESC
                LIMIT 1
            """, (func_id, data_norm, data_norm)).fetchone())
            if row:
                mult = row["multiplicador"]
        except Exception:
            pass

    conn.close()
    return mult


def _get_config_multiplicadores() -> dict:
    """Retorna configuração completa de multiplicadores para o painel admin."""
    conn = _conn()
    try:
        _criar_tabelas_multiplicadores(conn)
    except Exception:
        pass

    result = {
        "global": 1.5,
        "por_funcionario": [],
        "por_dia_semana": [],
        "por_periodo": [],
    }

    try:
        row = _d(_run(conn, "SELECT multiplicador FROM mult_config_global WHERE id=1").fetchone())
        if row:
            result["global"] = row["multiplicador"]
    except Exception:
        pass

    try:
        rows = [_d(r) for r in _run(conn, """
            SELECT f.id, f.nome, mc.multiplicador, mc.atualizado_em
            FROM mult_config_funcionario mc
            JOIN funcionarios f ON f.id = mc.funcionario_id
            ORDER BY f.nome
        """).fetchall()]
        result["por_funcionario"] = rows
    except Exception:
        pass

    try:
        rows = [_d(r) for r in _run(conn,
            "SELECT dia_semana, multiplicador, atualizado_em FROM mult_config_dia_semana ORDER BY dia_semana"
        ).fetchall()]
        result["por_dia_semana"] = rows
    except Exception:
        pass

    try:
        rows = [_d(r) for r in _run(conn, """
            SELECT mp.id, mp.funcionario_id, f.nome as func_nome,
                   mp.data_inicio, mp.data_fim, mp.multiplicador,
                   mp.descricao, mp.criado_em
            FROM mult_config_periodo mp
            LEFT JOIN funcionarios f ON f.id = mp.funcionario_id
            ORDER BY mp.data_inicio DESC
        """).fetchall()]
        result["por_periodo"] = rows
    except Exception:
        pass

    conn.close()
    return result


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

    # ── Calcula valor monetário ──────────────────────────────────────────────
    cur = _run(conn,
        "SELECT salario, tipo_pag, jornada_h FROM funcionarios WHERE id=?",
        (func_id,))
    func_sal = _d(cur.fetchone())
    valor_hora = 0.0
    valor_hora_ext = 0.0
    valor_a_receber = 0.0
    mult = 1.5
    if func_sal:
        sal  = func_sal.get("salario", 0) or 0
        tipo = func_sal.get("tipo_pag", "mes") or "mes"
        jorn = func_sal.get("jornada_h", 8) or 8
        horas_periodo = (jorn * 5) if tipo == "semana" else (jorn * 5 * 4.333)
        valor_hora = sal / horas_periodo if horas_periodo else 0
        mult = _get_multiplicador(func_id)
        valor_hora_ext = valor_hora * mult
        valor_a_receber = (saldo_min / 60.0) * valor_hora_ext if saldo_min > 0 else 0
    # total pago
    try:
        cur = _run(conn,
            "SELECT COALESCE(SUM(valor_pago),0) as total FROM pagamentos_horas WHERE funcionario_id=?",
            (func_id,))
        pago_row = _d(cur.fetchone())
        total_pago = pago_row["total"] if pago_row else 0
    except Exception:
        total_pago = 0
    a_pagar = max(0, valor_a_receber - total_pago)

    conn.close()

    def _brl(v):
        return "R$ {:,.2f}".format(v).replace(",","X").replace(".",",").replace("X",".")

    return {
        "id":               func_id,
        "nome":             nome,
        "saldo_min":        saldo_min,
        "saldo_fmt":        _saldo_fmt(saldo_min),
        "saldo_horas":      abs(saldo_min) // 60,
        "saldo_mins":       abs(saldo_min) % 60,
        "saldo_positivo":   saldo_min >= 0,
        "atualizado_em":    (atualizado or "")[:16],
        "historico":        historico,
        "valor_hora":       round(valor_hora, 4),
        "valor_hora_fmt":   _brl(valor_hora),
        "valor_hora_ext":   round(valor_hora_ext, 4),
        "valor_hora_ext_fmt": _brl(valor_hora_ext),
        "multiplicador":    mult,
        "valor_a_receber":  round(valor_a_receber, 2),
        "valor_fmt":        _brl(valor_a_receber),
        "total_pago":       round(total_pago, 2),
        "total_pago_fmt":   _brl(total_pago),
        "a_pagar":          round(a_pagar, 2),
        "a_pagar_fmt":      _brl(a_pagar),
    }


def buscar_registros(func_id: int, limite: int = 30) -> list:
    """Retorna registros diários do funcionário com horários, extras e valor."""
    conn = _conn()

    # Busca dados salariais
    cur = _run(conn,
        "SELECT salario, tipo_pag, jornada_h FROM funcionarios WHERE id=?",
        (func_id,))
    func_sal = _d(cur.fetchone())
    valor_hora = 0.0
    if func_sal:
        sal  = func_sal.get("salario", 0) or 0
        tipo = func_sal.get("tipo_pag", "mes") or "mes"
        jorn = func_sal.get("jornada_h", 8) or 8
        horas_periodo = (jorn * 5) if tipo == "semana" else (jorn * 5 * 4.333)
        valor_hora = sal / horas_periodo if horas_periodo else 0
        jornada_h = jorn
    else:
        jornada_h = 8.0

    jornada_min = int(jornada_h * 60)

    # Busca batidas mais recentes sem erro
    cur = _run(conn, """
        SELECT data, horarios, qtd_batidas, erros, tem_erro
        FROM batidas
        WHERE funcionario_id=? AND tem_erro=0
        ORDER BY data DESC
        LIMIT ?
    """, (func_id, limite))
    rows = [_d(r) for r in cur.fetchall()]
    conn.close()

    def _hm(t):
        """Converte 'HH:MM' para minutos desde meia-noite."""
        try:
            h, m = t.split(":")
            return int(h) * 60 + int(m)
        except Exception:
            return 0

    registros = []
    for r in rows:
        horarios_raw = r.get("horarios", "") or ""
        horarios = [h.strip() for h in horarios_raw.split(",") if h.strip()]
        qtd = len(horarios)

        trabalhados = 0
        if qtd == 2:
            trabalhados = max(0, _hm(horarios[1]) - _hm(horarios[0]))
        elif qtd >= 4:
            trabalhados = max(0, _hm(horarios[1]) - _hm(horarios[0]))
            trabalhados += max(0, _hm(horarios[3]) - _hm(horarios[2]))
        elif qtd > 0:
            for i in range(0, qtd - 1, 2):
                trabalhados += max(0, _hm(horarios[i+1]) - _hm(horarios[i]))

        saldo    = trabalhados - jornada_min
        extras   = max(0,  saldo)
        faltas   = max(0, -saldo)
        mult_dia = _get_multiplicador(func_id, r.get("data", ""))
        valor_hora_ext = valor_hora * mult_dia
        valor_ex = round((extras / 60.0) * valor_hora_ext, 2)

        def _fmt_min(m):
            return f"{m // 60:02d}:{m % 60:02d}"

        registros.append({
            "data":            r["data"],
            "horarios":        horarios,
            "trabalhados_fmt": _fmt_min(trabalhados),
            "trabalhados_min": trabalhados,
            "jornada_fmt":     _fmt_min(jornada_min),
            "extras_min":      extras,
            "extras_fmt":      _fmt_min(extras),
            "faltas_min":      faltas,
            "faltas_fmt":      _fmt_min(faltas),
            "saldo_min":       saldo,
            "saldo_fmt":       f"{'+' if saldo >= 0 else '-'}{_fmt_min(abs(saldo))}",
            "status":          "extra" if extras > 0 else ("falta" if faltas > 0 else "ok"),
            "valor_extras":    valor_ex,
            "valor_extras_fmt": "R$ {:,.2f}".format(valor_ex).replace(",","X").replace(".",",").replace("X","."),
            "multiplicador":   mult_dia,
        })

    return registros



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


@app.route("/api/registros")
@requer_token
def api_registros():
    """Retorna registros diários com horários, extras e valor monetário."""
    limite = min(int(request.args.get("limite", 30)), 90)
    return jsonify(buscar_registros(request.func_id, limite)), 200


@app.route("/api/registros-por-nome/<nome>")
def api_registros_por_nome(nome):
    """Registros diários por nome (sem token, mesmo padrão do saldo-por-nome)."""
    func = buscar_func_por_nome(nome)
    if not func:
        return jsonify({"erro": "Funcionário não encontrado."}), 404
    limite = min(int(request.args.get("limite", 30)), 90)
    return jsonify(buscar_registros(func["id"], limite)), 200


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



# ──────────────────────────────────────────────────────────────────────────────
# ROTAS DE MULTIPLICADOR (admin)
# ──────────────────────────────────────────────────────────────────────────────

def _check_admin_key():
    """Retorna True se a chave admin está OK (ou se não há chave configurada)."""
    if ADMIN_KEY:
        return request.headers.get("X-Admin-Key", "") == ADMIN_KEY
    return True


@app.route("/api/admin/multiplicadores", methods=["GET"])
def api_get_multiplicadores():
    if not _check_admin_key():
        return jsonify({"erro": "Não autorizado."}), 401
    return jsonify(_get_config_multiplicadores()), 200


@app.route("/api/admin/multiplicadores/global", methods=["POST"])
def api_set_mult_global():
    if not _check_admin_key():
        return jsonify({"erro": "Não autorizado."}), 401
    dados = request.get_json(silent=True) or {}
    try:
        mult = float(dados.get("multiplicador", 1.5))
        if mult <= 0 or mult > 10:
            return jsonify({"erro": "Multiplicador inválido (0 < mult <= 10)."}), 400
    except (ValueError, TypeError):
        return jsonify({"erro": "Multiplicador deve ser um número."}), 400
    conn = _conn()
    _criar_tabelas_multiplicadores(conn)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _run(conn,
        "INSERT INTO mult_config_global(id, multiplicador, atualizado_em) VALUES(1,?,?)"
        " ON CONFLICT(id) DO UPDATE SET multiplicador=excluded.multiplicador, atualizado_em=excluded.atualizado_em",
        (mult, now))
    conn.commit(); conn.close()
    return jsonify({"ok": True, "multiplicador": mult}), 200


@app.route("/api/admin/multiplicadores/funcionario", methods=["POST"])
def api_set_mult_funcionario():
    if not _check_admin_key():
        return jsonify({"erro": "Não autorizado."}), 401
    dados = request.get_json(silent=True) or {}
    func_id = dados.get("funcionario_id")
    mult_raw = dados.get("multiplicador")
    if not func_id:
        return jsonify({"erro": "funcionario_id obrigatório."}), 400
    try:
        mult = float(mult_raw)
        if mult <= 0 or mult > 10:
            return jsonify({"erro": "Multiplicador inválido."}), 400
    except (ValueError, TypeError):
        return jsonify({"erro": "Multiplicador deve ser um número."}), 400
    conn = _conn()
    _criar_tabelas_multiplicadores(conn)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _run(conn,
        "INSERT INTO mult_config_funcionario(funcionario_id, multiplicador, atualizado_em) VALUES(?,?,?)"
        " ON CONFLICT(funcionario_id) DO UPDATE SET multiplicador=excluded.multiplicador, atualizado_em=excluded.atualizado_em",
        (func_id, mult, now))
    conn.commit(); conn.close()
    return jsonify({"ok": True, "funcionario_id": func_id, "multiplicador": mult}), 200


@app.route("/api/admin/multiplicadores/funcionario/<int:func_id>", methods=["DELETE"])
def api_del_mult_funcionario(func_id):
    if not _check_admin_key():
        return jsonify({"erro": "Não autorizado."}), 401
    conn = _conn()
    _run(conn, "DELETE FROM mult_config_funcionario WHERE funcionario_id=?", (func_id,))
    conn.commit(); conn.close()
    return jsonify({"ok": True}), 200


@app.route("/api/admin/multiplicadores/dia-semana", methods=["POST"])
def api_set_mult_dia_semana():
    if not _check_admin_key():
        return jsonify({"erro": "Não autorizado."}), 401
    dados = request.get_json(silent=True) or {}
    dia = dados.get("dia_semana")
    mult_raw = dados.get("multiplicador")
    if dia is None or not (0 <= int(dia) <= 6):
        return jsonify({"erro": "dia_semana deve ser 0 (Dom) a 6 (Sab)."}), 400
    try:
        mult = float(mult_raw)
        if mult <= 0 or mult > 10:
            return jsonify({"erro": "Multiplicador inválido."}), 400
    except (ValueError, TypeError):
        return jsonify({"erro": "Multiplicador deve ser um número."}), 400
    conn = _conn()
    _criar_tabelas_multiplicadores(conn)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _run(conn,
        "INSERT INTO mult_config_dia_semana(dia_semana, multiplicador, atualizado_em) VALUES(?,?,?)"
        " ON CONFLICT(dia_semana) DO UPDATE SET multiplicador=excluded.multiplicador, atualizado_em=excluded.atualizado_em",
        (int(dia), mult, now))
    conn.commit(); conn.close()
    return jsonify({"ok": True, "dia_semana": int(dia), "multiplicador": mult}), 200


@app.route("/api/admin/multiplicadores/dia-semana/<int:dia>", methods=["DELETE"])
def api_del_mult_dia_semana(dia):
    if not _check_admin_key():
        return jsonify({"erro": "Não autorizado."}), 401
    conn = _conn()
    _run(conn, "DELETE FROM mult_config_dia_semana WHERE dia_semana=?", (dia,))
    conn.commit(); conn.close()
    return jsonify({"ok": True}), 200


@app.route("/api/admin/multiplicadores/periodo", methods=["POST"])
def api_add_mult_periodo():
    if not _check_admin_key():
        return jsonify({"erro": "Não autorizado."}), 401
    dados = request.get_json(silent=True) or {}
    data_inicio = dados.get("data_inicio", "").strip()
    data_fim    = dados.get("data_fim", "").strip()
    mult_raw    = dados.get("multiplicador")
    func_id     = dados.get("funcionario_id")
    descricao   = dados.get("descricao", "").strip()
    if not data_inicio or not data_fim:
        return jsonify({"erro": "data_inicio e data_fim obrigatorios (YYYY-MM-DD)."}), 400
    try:
        mult = float(mult_raw)
        if mult <= 0 or mult > 10:
            return jsonify({"erro": "Multiplicador inválido."}), 400
    except (ValueError, TypeError):
        return jsonify({"erro": "Multiplicador deve ser um número."}), 400
    conn = _conn()
    _criar_tabelas_multiplicadores(conn)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur = _run(conn,
        "INSERT INTO mult_config_periodo(funcionario_id,data_inicio,data_fim,multiplicador,descricao,criado_em)"
        " VALUES(?,?,?,?,?,?)",
        (func_id, data_inicio, data_fim, mult, descricao, now))
    new_id = cur.lastrowid
    conn.commit(); conn.close()
    return jsonify({"ok": True, "id": new_id}), 200


@app.route("/api/admin/multiplicadores/periodo/<int:pid>", methods=["DELETE"])
def api_del_mult_periodo(pid):
    if not _check_admin_key():
        return jsonify({"erro": "Não autorizado."}), 401
    conn = _conn()
    _run(conn, "DELETE FROM mult_config_periodo WHERE id=?", (pid,))
    conn.commit(); conn.close()
    return jsonify({"ok": True}), 200


@app.route("/api/admin/multiplicadores/auditoria", methods=["POST"])
def api_registrar_auditoria():
    if not _check_admin_key():
        return jsonify({"erro": "Não autorizado."}), 401
    dados = request.get_json(silent=True) or {}
    func_id    = dados.get("funcionario_id")
    referencia = dados.get("referencia", "").strip()
    mult       = dados.get("multiplicador")
    valor_hora = dados.get("valor_hora", 0)
    saldo_min  = dados.get("saldo_min", 0)
    obs        = dados.get("observacao", "").strip()
    if not func_id or not referencia or mult is None:
        return jsonify({"erro": "funcionario_id, referencia e multiplicador obrigatorios."}), 400
    try:
        mult = float(mult); valor_hora = float(valor_hora); saldo_min = int(saldo_min)
    except (ValueError, TypeError):
        return jsonify({"erro": "Valores inválidos."}), 400
    valor_hora_ext = valor_hora * mult
    valor_total    = (saldo_min / 60.0) * valor_hora_ext if saldo_min > 0 else 0
    conn = _conn()
    _criar_tabelas_multiplicadores(conn)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _run(conn,
        "INSERT INTO mult_auditoria(funcionario_id,referencia,multiplicador,valor_hora,valor_hora_ext,saldo_min,valor_total,criado_em,observacao)"
        " VALUES(?,?,?,?,?,?,?,?,?)",
        (func_id, referencia, mult, valor_hora, valor_hora_ext, saldo_min, valor_total, now, obs))
    conn.commit(); conn.close()
    return jsonify({"ok": True, "valor_total": round(valor_total, 2)}), 200


@app.route("/api/admin/multiplicadores/auditoria", methods=["GET"])
def api_get_auditoria():
    if not _check_admin_key():
        return jsonify({"erro": "Não autorizado."}), 401
    func_id = request.args.get("funcionario_id")
    conn = _conn()
    _criar_tabelas_multiplicadores(conn)
    if func_id:
        rows = [_d(r) for r in _run(conn, """
            SELECT ma.*, f.nome as func_nome
            FROM mult_auditoria ma JOIN funcionarios f ON f.id = ma.funcionario_id
            WHERE ma.funcionario_id=? ORDER BY ma.criado_em DESC LIMIT 100
        """, (int(func_id),)).fetchall()]
    else:
        rows = [_d(r) for r in _run(conn, """
            SELECT ma.*, f.nome as func_nome
            FROM mult_auditoria ma JOIN funcionarios f ON f.id = ma.funcionario_id
            ORDER BY ma.criado_em DESC LIMIT 100
        """).fetchall()]
    conn.close()
    return jsonify(rows), 200


@app.route("/api/multiplicador-efetivo/<nome>")
def api_mult_efetivo_por_nome(nome):
    func = buscar_func_por_nome(nome)
    if not func:
        return jsonify({"erro": "Funcionário não encontrado."}), 404
    data = request.args.get("data")
    mult = _get_multiplicador(func["id"], data)
    return jsonify({"nome": nome, "multiplicador": mult, "data": data}), 200


@app.route("/api/health")
def api_health():
    """Healthcheck — Railway verifica se o serviço está no ar."""
    return jsonify({
        "status": "ok",
        "banco":  "postgresql" if USE_POSTGRES else "sqlite",
        "hora":   datetime.now().isoformat(),
    }), 200


@app.route("/api/admin/sincronizar", methods=["POST"])
def api_admin_sincronizar():
    """
    Recebe saldo e movimentos de todos os funcionários enviados pelo ponto.py
    e atualiza o banco da nuvem. Protegido pela ADMIN_KEY (se configurada).
    """
    if ADMIN_KEY:
        if request.headers.get("X-Admin-Key", "") != ADMIN_KEY:
            return jsonify({"erro": "Não autorizado. Verifique a ADMIN_KEY."}), 401

    dados = request.get_json(silent=True) or {}
    funcionarios = dados.get("funcionarios", [])
    if not funcionarios:
        return jsonify({"erro": "Nenhum dado recebido."}), 400

    conn = _conn()
    now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    atualizados = 0

    try:
        for func in funcionarios:
            nome      = (func.get("nome") or "").strip()
            saldo_min = int(func.get("saldo_min", 0))
            pin       = func.get("pin") or None
            movs      = func.get("movimentos", [])
            if not nome:
                continue

            # ── Upsert funcionário ──────────────────────────────────────────
            _run(conn,
                "INSERT INTO funcionarios(nome,pin,salario,tipo_pag,jornada_h,criado_em,ativo)"
                " VALUES(?,?,0,'mes',8,?,1)"
                " ON CONFLICT(nome) DO UPDATE SET pin=excluded.pin",
                (nome, pin, now))

            cur = _run(conn, "SELECT id FROM funcionarios WHERE nome=?", (nome,))
            row = _d(cur.fetchone())
            if not row:
                continue
            fid = row["id"]

            # ── Upsert saldo ────────────────────────────────────────────────
            _run(conn,
                "INSERT INTO banco_horas(funcionario_id,saldo_min,atualizado_em)"
                " VALUES(?,?,?)"
                " ON CONFLICT(funcionario_id) DO UPDATE"
                " SET saldo_min=excluded.saldo_min, atualizado_em=excluded.atualizado_em",
                (fid, saldo_min, now))

            # ── Insere movimentos novos (evita duplicatas por criado_em+tipo+minutos) ──
            for mov in movs:
                criado_em = (mov.get("criado_em") or now)[:19]
                cur2 = _run(conn,
                    "SELECT id FROM banco_horas_mov"
                    " WHERE funcionario_id=? AND tipo=? AND minutos=? AND criado_em=?",
                    (fid, mov.get("tipo","credito"), int(mov.get("minutos",0)), criado_em))
                if not _d(cur2.fetchone()):
                    _run(conn,
                        "INSERT INTO banco_horas_mov"
                        "(funcionario_id,tipo,minutos,descricao,referencia,criado_em)"
                        " VALUES(?,?,?,?,?,?)",
                        (fid, mov.get("tipo","credito"), int(mov.get("minutos",0)),
                         mov.get("descricao",""), mov.get("referencia",""), criado_em))

            atualizados += 1

        conn.commit()
        conn.close()
        return jsonify({"ok": True, "atualizados": atualizados}), 200

    except Exception as e:
        try:
            conn.rollback()
            conn.close()
        except Exception:
            pass
        return jsonify({"erro": str(e)}), 500


# ──────────────────────────────────────────────────────────────────────────────
# SERVE O APP (PWA)
# ──────────────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    resp = send_from_directory(STATIC_DIR, "index.html")
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


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
